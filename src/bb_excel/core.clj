(ns bb-excel.core
  (:require [bb-excel.util :refer [by-tag find-first throw-ex]]
            [clojure.data.xml :as xml]
            [hiccup2.core :as hc]
            [clojure.java.io :as io]
            [clojure.set :refer [rename-keys]]
            [clojure.string :as str])
  (:import [java.io File FileOutputStream]
           [java.text SimpleDateFormat]
           [java.time LocalDate LocalDateTime Month]
           [java.time.format DateTimeFormatter]
           [java.time.temporal ChronoUnit]
           [java.util TimeZone]
           [java.util.zip ZipEntry ZipFile ZipOutputStream])
  (:gen-class))

(set! *warn-on-reflection* true)

(def  ^SimpleDateFormat SDF (doto (SimpleDateFormat. "HH:mm:ss")
                              (.setTimeZone (TimeZone/getTimeZone "UTC"))))
(def ^:const BASE_ROW_INDEX 0)
(def ^:const BASE_COLUMN_INDEX 0)
(def ^:const A_CHAR_INDEX (int \A))

(def ^:const dates #{"14"  "15"  "16"  "17"  "30"  "34"  "51"
                     "52"  "53"  "55"  "56"  "58"  "165"
                     "166" "167" "168" "169" "170" "171" "172"
                     "173" "174" "175" "176" "177" "178" "179"
                     "180" "181" "184" "185" "186" "187"})

(def ^:const times #{"164"  "18" "19" "21" "20"  "45" "46" "47"})

(def ^:const pcts  #{"9" "10"})

(def ^:const error-codes
  {"#NAME?"   :bad-name
   "#DIV/0!"  :div-by-0
   "#REF!"    :invalid-reference
   "#NUM!"    :infinity
   "#N/A"     :not-applicable
   "#VALUE!"  :invalid-value
   "#NULL!"   :null
   "#SPILL!"  :multiple-results
   nil        :unknown-error})

(def ^:const defaults
  "Default values for processing the Excel Spreadsheet
   :row  integer  :-  Which row to begin data extraction defaults to 0 
   :fxn  function :-  Which function to use parse header rows
   :rows integer  :-  Number of rows to extract
   :hdr  boolean  :- Rename columns with data from the first row"
  {:row 0
   :fxn str
   :rows 10000
   :hdr false})

(def xmlh "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")

(def xlns {:xmlns       "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           :xmlns:r     "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           :xmlns:mx    "http://schemas.microsoft.com/office/mac/excel/2008/main"
           :xmlns:mc    "http://schemas.openxmlformats.org/markup-compatibility/2006"
           :xmlns:mv    "urn:schemas-microsoft-com:mac:vml"
           :xmlns:x14   "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
           :xmlns:x15   "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
           :xmlns:x14ac "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
           :xmlns:xm    "http://schemas.microsoft.com/office/excel/2006/main"})

(def wb-relationships
  (str xmlh
       (hc/html [:Relationships {:xmlns "http://schemas.openxmlformats.org/package/2006/relationships"}
                 [:Relationship {:Id "rId1"
                                 :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                                 :Target "xl/workbook.xml"}]])))

(defn parse-xlong
  "Adds support for passing in nil into parse-long
   Return an empty string if x is nil"
  [x]
  (parse-long (or x "")))

(defn- get-zipfile
  "Retrieve ZipFile object if provided `file-or-filename` point to existing file."
  [file-or-filename]
  (when-let [^File file (condp instance? file-or-filename
                          String (io/file file-or-filename)
                          File file-or-filename
                          (throw-ex (format "Could not open '%s'! Argument should be string or file." file-or-filename)))]
    (if (.exists file)
      (ZipFile. file)
      (throw-ex (format "Could not open '%s'! File does not exist." file-or-filename)))))

(defn get-workbook-relationships
  "Get the relationship mappings from xl/_rels/workbook.xml.rels.
   Returns a map from rId to the Target path (e.g. {\"rId1\" \"worksheets/sheet1.xml\"})."
  [^ZipFile zipfile]
  (if-let [rels-entry (.getEntry zipfile "xl/_rels/workbook.xml.rels")]
    (with-open [rels (.getInputStream zipfile rels-entry)]
      (let [rels-node (xml/parse rels {:namespace-aware false})
            rel-nodes (->> (:content rels-node)
                           (filter (by-tag :Relationship)))]
        (into {} (map (fn [rel]
                        (let [attrs (:attrs rel)]
                          [(:Id attrs) (:Target attrs)])))
              rel-nodes)))
    {}))

(defn get-sheet-names*
  [^ZipFile zipfile]
  (if-let [workbook-entry (.getEntry zipfile "xl/workbook.xml")]
    (with-open [workbook (.getInputStream zipfile workbook-entry)]
      (let [workbook-node (xml/parse workbook {:namespace-aware false})
            sheets-node (->> (:content workbook-node)
                             (find-first (by-tag :sheets)))
            sheet-nodes (->> (:content sheets-node)
                             (filter (by-tag :sheet)))]
        (into [] (comp (map :attrs)
                       (map #(select-keys % [:sheetId :name :id]))
                       (map #(update % :sheetId parse-xlong))
                       (map #(rename-keys % {:sheetId :idx :id :rid})))
              sheet-nodes)))
    []))

(defn get-sheet-names
  "Retrieves a list of Sheet Names from a given Excel Spreadsheet.
   Returns a vector of maps with :name and :idx keys."
  [file-or-filename]
  (let [^ZipFile zipfile (get-zipfile file-or-filename)]
    (mapv #(dissoc % :rid) (get-sheet-names* zipfile))))

(defn num2date
  "Format Excel Date"
  [n]
  (when n (.format (.plusDays (LocalDate/of 1899 Month/DECEMBER 30)  (parse-double (str n)))
                   (DateTimeFormatter/ofPattern "MM/dd/yyyy"))))

(defn num2time
  "Format Excel Time"
  [n]
  (when n (.format SDF (*  (parse-double (str n)) 24 60 60 1000))))

(defn num2pct
  "Format Percentage"
  [n]
  (when n (format "%.4f%%" (* 100 (parse-double (str n))))))

(defn style-check
  "Check if the style id is within a range."
  [cell-attrs styles ids]
  (when (:s cell-attrs)
    (try
      (ids (styles (parse-xlong (:s cell-attrs))))
      (catch Exception _ false))))

(defn extract-cell-value
  "Possible cell-value types well explained here https://stackoverflow.com/a/18346273"
  [shared-strings styles cell]
  (let [raw-cell-value (-> cell :content last :content last)
        cell-attrs (:attrs cell)
        cell-type (:t cell-attrs)]
    (cond
      (= cell-type "s")                 (nth shared-strings (parse-xlong raw-cell-value))
      (= cell-type "str")               raw-cell-value
      (= cell-type "inlineStr")         (-> raw-cell-value :content last)
      (= cell-type "b")                 (if (= "1" raw-cell-value) true false)
      (= cell-type "e")                 (get error-codes raw-cell-value)
      (= cell-type "n")                 (parse-xlong raw-cell-value)
      (style-check cell-attrs styles pcts)    (num2pct raw-cell-value)
      (style-check cell-attrs styles dates)   (num2date raw-cell-value)
      (style-check cell-attrs styles times)   (num2time raw-cell-value)
      :else raw-cell-value)))

(defn- get-cell-text
  "Extract text from cell"
  [cell]
  (->> (xml-seq cell)
       (filter (by-tag :t))
       (mapcat :content)
       (str/join)))

(defn get-shared-strings
  "Get dictionary of all unique strings in the Excel spreadsheet"
  [^ZipFile zipfile]
  (if-let [shared-strings-entry (.getEntry zipfile "xl/sharedStrings.xml")]
    (with-open [shared-strings (.getInputStream zipfile shared-strings-entry)]
      (let [sst-node (xml/parse shared-strings {:namespace-aware false})]
        (mapv get-cell-text (:content sst-node))))
    []))

(defn get-styles
  [^ZipFile zipfile]
  (if-let [styles-entry (.getEntry zipfile "xl/styles.xml")]
    (with-open [styles (.getInputStream zipfile styles-entry)]
      (let [style-sheet-node (xml/parse styles {:namespace-aware false})
            cell-xfs-node (->> (:content style-sheet-node)
                               (find-first (by-tag :cellXfs)))
            xf-nodes (->> (:content cell-xfs-node)
                          (filter (by-tag :xf)))]
        (mapv #(-> % :attrs :numFmtId) xf-nodes)))
    []))

(defn valid-cell-index?
  [cell-index]
  (if cell-index
    (boolean (re-find #"^[A-Z]{1,3}\d+$" cell-index))
    false))

(defn number->column-letter
  [n]
  (loop [num n
         acc ""]
    (if (> num 0)
      (let [residue (mod (dec num) 26)
            new-num (quot (dec num) 26)]
        (recur new-num (str (char (+ residue A_CHAR_INDEX)) acc)))
      acc)))

(defn column-letter->number
  "Convert column letter(s) to a 1-based numeric index.
   A=1, B=2, ..., Z=26, AA=27, etc."
  [col-str]
  (reduce (fn [acc c]
            (+ (* acc 26) (- (int c) (dec A_CHAR_INDEX))))
          0
          col-str))

(defn get-col-index
  "Returns a vector of [col-letter col-number] where col-number is the 1-based numeric index.
   Self-calculated index is used only if cell-index attribute(:r) is missing on the cell"
  [cell last-processed-col-number]
  (let [cell-index (-> cell :attrs :r)]
    (if (valid-cell-index? cell-index)
      (let [col-letter (re-find #"[A-Z]{1,3}" cell-index)]
        [col-letter (column-letter->number col-letter)])
      (let [new-col-number (inc last-processed-col-number)]
        [(number->column-letter new-col-number) new-col-number]))))

(defn process-row
  "Process Excel row of data"
  [shared-strings styles row]
  (->> (:content row)
       (reduce (fn [{:keys [row-data last-processed-col-number]} cell]
                 (let [[col-letter col-number] (get-col-index cell last-processed-col-number)
                       cell-value (extract-cell-value shared-strings styles cell)]
                   {:row-data (assoc row-data (keyword col-letter) cell-value)
                    :last-processed-col-number col-number}))
               {:row-data {}
                :last-processed-col-number BASE_COLUMN_INDEX})
       (:row-data)))

(defn process-rows
  [shared-strings styles last-processed-row-index rows]
  (lazy-seq
   (when rows
     (let [row (first rows)
           row-index (or (some-> row :attrs :r parse-xlong)
                         (inc last-processed-row-index))
           processed-row (process-row shared-strings styles row)]
       (cons (assoc processed-row
                    :_r row-index)
             (process-rows shared-strings
                           styles
                           row-index
                           (next rows)))))))

(defn find-sheet-by-name-or-index
  "Find a sheet by name (case-insensitive) or by positional index (1-based).
   When using an integer index, it refers to the position in the sheets list,
   not the internal sheetId."
  [sheets sheetname-or-idx]
  (cond
    (string? sheetname-or-idx)
    (find-first (fn [sheet]
                  (= (str/lower-case sheetname-or-idx)
                     (str/lower-case (:name sheet))))
                sheets)

    (and (integer? sheetname-or-idx) (pos? sheetname-or-idx))
    ;; Use 1-based positional index, not sheetId
    (nth sheets (dec sheetname-or-idx) nil)

    :else nil))

(defn get-sheet-entry
  "Get the ZipEntry for a worksheet using the relationship ID.
   The rels map provides the mapping from rId to the actual worksheet path."
  [^ZipFile zipfile rels rid]
  (if-let [target (get rels rid)]
    (let [path (if (str/starts-with? target "/")
                 (subs target 1)
                 (str "xl/" target))]
      (or (.getEntry zipfile path)
          (throw-ex (format "Could not find worksheet file '%s' for relationship '%s'!" path rid))))
    (throw-ex (format "Could not find relationship with id '%s'!" rid))))

(defn get-sheet
  "Get sheet from file or filename"
  ([file-or-filename]
   (get-sheet file-or-filename 1 {}))
  ([file-or-filename sheetname-or-idx]
   (get-sheet file-or-filename sheetname-or-idx {}))
  ([file-or-filename sheetname-or-idx options]
   (let [^ZipFile zipfile (get-zipfile file-or-filename)
         sheets (get-sheet-names* zipfile)
         found-sheet (find-sheet-by-name-or-index sheets sheetname-or-idx)
         _ (when-not found-sheet
             (throw-ex (format "Could not find sheet with name or index equal '%s'! Sheet does not exist." sheetname-or-idx)))
         rels (get-workbook-relationships zipfile)
         ^ZipEntry sheet-entry (get-sheet-entry zipfile rels (:rid found-sheet))
         opts    (merge defaults options)
         row     (:row opts)
         hdr     (:hdr opts)
         row     (if (and hdr (zero? row)) 1 row)
         rows    (:rows opts)
         fxn     (:fxn opts)
         cols    (map fxn (:columns opts))
         shared-strings (get-shared-strings zipfile)
         styles  (get-styles zipfile)]
     (with-open [sheet (.getInputStream zipfile sheet-entry)]
       (let [worksheet-node (xml/parse sheet {:namespace-aware false})
             sheet-data-node (->> (:content worksheet-node)
                                  (find-first (by-tag :sheetData)))
             row-nodes (:content sheet-data-node)
             d (->> row-nodes
                    (take rows)
                    (process-rows shared-strings
                                  styles
                                  BASE_ROW_INDEX))
             dx (remove #(= row (:_r %)) d)
             h (when hdr (merge (update-vals (first (filter #(= (:_r %) row) d)) fxn)
                                {:_r :_r}))
             dy (if (pos? rows)
                  (take rows (mapv #(rename-keys % h) dx))
                  (mapv #(rename-keys % h) dx))]
         (if (empty? cols) dy (mapv #(select-keys % cols) dy)))))))

(defn get-sheets
  "Get all or specified sheet from the excel spreadsheet"
  ([file-or-filename]
   (get-sheets file-or-filename {}))
  ([file-or-filename options]
   (let [sns  (get-sheet-names file-or-filename)
         sxs  (if (:sheet options) (filter #(= (:sheet options) (:name %)) sns) sns)
         res  (if (empty? sxs) [{:sheet []}]
                  (map #(assoc % :sheet
                               (try (get-sheet file-or-filename (:name %) options)
                                    (catch Exception ex [(bean ex)]))) sxs))]
     res)))

(defn when-num
  "Returns nil for empty strings when a number is expected"
  [s]
  (cond
    (empty? s) nil
    (number? (read-string s))
    (Integer/parseInt s)
    :else 0))

(defn when-str
  "Returns nil for empty strings"
  [s]
  (cond
    (empty? s) nil
    :else s))

(defn parse-range
  "Takes in an Excel coordinate and returns a hashmap of rows and columns to pull"
  [s]
  (let [[_ osc osr oec oer] (re-matches #"([A-Z]+)([0-9]*)[:]?([A-Z]*)([0-9]*)" s)
        sc (or osc "A")
        ec (or (when-str oec) (when-str osc) sc)
        sr (or (when-num osr) 1)
        er (or (when-num oer) (when-num osr) 10000)]
    {:cols [sc ec]
     :rows [sr (inc er)]}))

(defn to-col
  "Takes in an ordinal and returns its equivalent column heading."
  [num]
  (loop [n num s ()]
    (if (> n 25)
      (let [r (mod n 26)]
        (recur (dec (/ (- n r) 26)) (cons (char (+ 65 r)) s)))
      (keyword (apply str (cons (char (+ 65 n)) s))))))

(defn crange
  "Creates as sequence of columns given a starting and ending column name."
  [s e]
  (cons :_r (let [sn (reduce + (map * (iterate (partial * 26) 1)
                                    (reverse (map (comp (partial + -64) int identity) s))))
                  en  (reduce + (map * (iterate (partial * 26) 1)
                                     (reverse (map (comp (partial + -64) int identity) e))))]
              (map to-col (range (dec sn) en)))))

(defn get-row
  "Get row from sheet by row index"
  [sheet row]
  (first (filter #(= row (:_r %)) sheet)))

(defn get-col
  "Get column from sheet by name. 
   If columns have been renamed use the new name."
  [sheet col]
  (map #(select-keys % [:_r col]) sheet))

(defn get-cells
  "Get range of values returned as list of rows"
  [sheet rows cols]
  (map #(select-keys % cols)
       (filter #(contains? (set rows) (:_r %)) sheet)))

(defn get-range
  "Get range of values using Excel cell coordinates
   e.g A1:C5"
  [sheet rg]
  (let [{:keys [cols rows]} (parse-range rg)
        [rs re] rows
        [cs ce] cols]
    (get-cells sheet (range rs re) (crange cs ce))))

(defn- normalize-xl-path
  "Normalize a relationship Target to a path within the zip (no leading slash, starts from root)."
  [target]
  (when target
    (let [t (if (str/starts-with? target "/") (subs target 1) target)]
      (if (str/starts-with? t "xl/") t (str "xl/" t)))))

(defn- get-sheet-table-paths
  "Return zip-relative paths to table XML files referenced by a worksheet.
   sheet-rel-path is the path relative to xl/ (e.g. 'worksheets/sheet1.xml')."
  [^ZipFile zipfile sheet-rel-path]
  (let [sheet-name (last (str/split sheet-rel-path #"/"))
        rels-path (str "xl/worksheets/_rels/" sheet-name ".rels")
        table-type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"]
    (if-let [rels-entry (.getEntry zipfile rels-path)]
      (with-open [rels (.getInputStream zipfile rels-entry)]
        (let [rels-node (xml/parse rels {:namespace-aware false})]
          (->> (:content rels-node)
               (filter (by-tag :Relationship))
               (filter #(= table-type (-> % :attrs :Type)))
               (mapv (fn [rel]
                       (let [target (-> rel :attrs :Target)]
                         (cond
                           (str/starts-with? target "/") (subs target 1)
                           (str/starts-with? target "../") (str "xl/" (subs target 3))
                           :else (str "xl/worksheets/" target))))))))
      [])))

(defn- parse-table-xml
  "Parse a table XML file and return table metadata."
  [^ZipFile zipfile table-path]
  (when-let [table-entry (.getEntry zipfile table-path)]
    (with-open [tstream (.getInputStream zipfile table-entry)]
      (let [table-node (xml/parse tstream {:namespace-aware false})
            attrs (:attrs table-node)
            cols-node (->> (:content table-node)
                           (find-first (by-tag :tableColumns)))
            columns (->> (:content cols-node)
                         (filter (by-tag :tableColumn))
                         (mapv #(-> % :attrs :name)))]
        {:name (:name attrs)
         :display-name (:displayName attrs)
         :ref (:ref attrs)
         :columns columns}))))

(defn get-table-names
  "Get metadata for all named tables in an Excel file.
   Returns a vector of maps with :name, :display-name, :sheet, :columns, and :ref keys."
  [file-or-filename]
  (let [^ZipFile zipfile (get-zipfile file-or-filename)
        sheets (get-sheet-names* zipfile)
        rels (get-workbook-relationships zipfile)]
    (into []
          (for [sheet sheets
                :let [target (get rels (:rid sheet))
                      sheet-zip-path (normalize-xl-path target)
                      ;; Path relative to xl/ for finding the _rels file
                      sheet-rel-path (when sheet-zip-path (subs sheet-zip-path 3))
                      table-paths (if sheet-rel-path
                                    (get-sheet-table-paths zipfile sheet-rel-path)
                                    [])]
                table-path table-paths
                :let [table-meta (parse-table-xml zipfile table-path)]
                :when table-meta]
            (assoc table-meta :sheet (:name sheet))))))

(defn get-table
  "Get data from a named Excel table.
   Returns a map with :name and :data keys.
   :data is a vector of maps using the table's column names as string keys."
  [file-or-filename table-name]
  (let [tables (get-table-names file-or-filename)
        table (find-first #(= table-name (:name %)) tables)]
    (when-not table
      (throw-ex (format "Could not find table '%s'! Table does not exist." table-name)))
    (let [ref (:ref table)
          columns (:columns table)
          {:keys [cols rows]} (parse-range ref)
          [cs ce] cols
          [rs re] rows
          ;; Exclude header row: data starts at row (inc rs), ends at row (dec re)
          data-range (str cs (inc rs) ":" ce (dec re))
          ;; Map column letter keywords to column name strings
          start-col-num (column-letter->number cs)
          col-rename (into {} (map-indexed
                               (fn [i col-name]
                                 [(keyword (number->column-letter (+ start-col-num i))) col-name])
                               columns))
          sheet-data (get-sheet file-or-filename (:sheet table))
          range-data (get-range sheet-data data-range)]
      {:name table-name
       :data (mapv #(-> % (rename-keys col-rename) (dissoc :_r)) range-data)})))

(defn ws-relationships [n]
  (str xmlh
       (hc/html
        (into [:Relationships {:xmlns "http://schemas.openxmlformats.org/package/2006/relationships"}]
              (for [x (range n)]
                [:Relationship {:Id (str "rId" (inc x))
                                :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                                :Target (str "worksheets/sheet" (inc x) ".xml")}])))))

(defn- content-types
  "Generate Content Types, optionally including table content type overrides."
  ([n] (content-types n []))
  ([n table-ids]
   (str xmlh
        (hc/html
         (into [:Types {:xmlns "http://schemas.openxmlformats.org/package/2006/content-types"}
                [:Default {:Extension :rels
                           :ContentType "application/vnd.openxmlformats-package.relationships+xml"}]
                [:Default {:Extension :xml
                           :ContentType :application/xml}]
                [:Override {:PartName "/xl/workbook.xml"
                            :ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"}]]
               (concat
                (for [x (range n)]
                  [:Override {:PartName (str "/xl/worksheets/sheet" (inc x) ".xml")
                              :ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"}])
                (for [tid table-ids]
                  [:Override {:PartName (str "/xl/tables/table" tid ".xml")
                              :ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"}])))))))

(defn excel-date-serial
  "Convert a java LocalDate to an MS Excel integer value"
  [datetime]
  (.between ChronoUnit/DAYS (LocalDate/of 1899 Month/DECEMBER 30) datetime))

(defn excel-time-serial
  "Convert a java LocalDateTime to an MS Excel decimal value."
  [datetime]
  (/ (.between ChronoUnit/SECONDS (LocalDateTime/of 1899 Month/DECEMBER 30 0 0) datetime) 86400.0))

(defn- cell-type
  "Determine cell data type"
  [value]
  (cond
    (instance? java.time.LocalDate value) ["n" [:v (excel-date-serial value)]]
    (instance? java.time.LocalDateTime value) ["n" [:v (excel-time-serial value)]]
    (string? value) ["inlineStr" [:is [:t value]]]
    (number? value) ["n" [:v value]]
    (boolean? value) ["b" [:v (if value "1" "0")]]
    :else ["inlineStr" [:is [:t (str value)]]]))

(defn- generate-xml-cell
  "Generate cell value in hiccup format"
  [c r value]
  (let [[t v] (cell-type value)]
    [:c {:r (str (if (keyword? c) (name c) c)
                 (inc r))
         :t t} v]))

(defn- generate-xml-row
  "Generate row information in hiccup format"
  ([row-data row-num]
   [:row {:r (inc row-num)}
    (map-indexed (fn [col-num val]
                   (generate-xml-cell (char (+ col-num 65)) row-num val))
                 row-data)])
  ([row-data row-num column-mapping]
   [:row {:r (inc row-num)}
    (if (nil? column-mapping)
      (map-indexed (fn [col-num val]
                     (generate-xml-cell (char (+ col-num 65)) row-num val))
                   (vals row-data))
      (for [[key val] row-data
            :let [col-letter (column-mapping key)]
            :when col-letter]
        (generate-xml-cell col-letter row-num val)))]))

(defn- create-sheet-xml
  "Create the sheet data in hiccup format. 
   Checks to see if the data provided is a vector of hashmaps vs a vector of vectors"
  [data]
  (let [headers (if (map? (first (:sheet data)))
                  (keys (first (:sheet data)))
                  (first (:sheet data)))
        rows (if (map? (first (:sheet data)))
               (map-indexed #(generate-xml-row %2 (inc %) (:cmap data)) (:sheet data))
               (map-indexed #(generate-xml-row %2 (inc %)) (rest (:sheet data))))]
    (str (hc/html [:worksheet xlns
                   [:sheetData (cons (generate-xml-row headers 0) rows)]]))))

(defn- extract-table-headers
  "Return a vector of string header names from the first row of sheet data."
  [sheet-data]
  (if (map? (first sheet-data))
    (mapv #(if (keyword? %) (name %) (str %)) (keys (first sheet-data)))
    (mapv str (first sheet-data))))

(defn- table-row-count
  "Return the number of data rows (excluding the header row)."
  [sheet-data]
  (if (map? (first sheet-data))
    (count sheet-data)
    (dec (count sheet-data))))

(defn- make-table-ref
  "Build an Excel range string covering header + data rows, e.g. 'A1:C4'."
  [col-count row-count]
  (str "A1:" (number->column-letter col-count) (inc row-count)))

(defn- sanitize-table-name
  "Replace characters that Excel forbids in table names with underscores."
  [s]
  (let [cleaned (str/replace (str s) #"[^A-Za-z0-9_\.]" "_")]
    (if (re-matches #"[0-9].*" cleaned) (str "T" cleaned) cleaned)))

(defn- make-table-xml
  "Generate the xl/tables/tableN.xml content for an Excel table."
  [table-id table-name table-style headers row-count]
  (let [ref (make-table-ref (count headers) row-count)]
    (str xmlh
         (hc/html
          [:table {:xmlns "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                   :id table-id
                   :name table-name
                   :displayName table-name
                   :ref ref}
           [:autoFilter {:ref ref}]
           (into [:tableColumns {:count (count headers)}]
                 (map-indexed (fn [i h] [:tableColumn {:id (inc i) :name h}]) headers))
           [:tableStyleInfo {:name table-style
                             :showFirstColumn "0"
                             :showLastColumn "0"
                             :showRowStripes "1"
                             :showColumnStripes "0"}]]))))

(defn- make-sheet-rels-xml
  "Generate xl/worksheets/_rels/sheetN.xml.rels pointing to a table file."
  [table-id]
  (str xmlh
       (hc/html
        [:Relationships {:xmlns "http://schemas.openxmlformats.org/package/2006/relationships"}
         [:Relationship {:Id "rId1"
                         :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"
                         :Target (str "../tables/table" table-id ".xml")}]])))

(defn- create-table-sheet-xml
  "Create worksheet XML that includes a <tableParts> reference to the table."
  [data]
  (let [sheet-data (:sheet data)
        headers    (extract-table-headers sheet-data)
        rows       (if (map? (first sheet-data))
                     (map-indexed #(generate-xml-row %2 (inc %) (:cmap data)) sheet-data)
                     (map-indexed #(generate-xml-row %2 (inc %)) (rest sheet-data)))]
    (str (hc/html [:worksheet xlns
                   [:sheetData (cons (generate-xml-row headers 0) rows)]
                   [:tableParts {:count 1}
                    [:tablePart {:r:id "rId1"}]]]))))

(defn create-zip-entry
  "For a given filepath and content add to a java ZipOuputStream"
  [zip-stream entry-name content]
  (let [entry  (ZipEntry. ^String entry-name)]
    (.putNextEntry ^ZipOutputStream zip-stream ^ZipEntry entry)
    (.write ^ZipOutputStream zip-stream (.getBytes ^String content "UTF-8"))
    (.closeEntry ^ZipOutputStream zip-stream)))

(defn create-xlsx
  "Create an Excel spreadsheet.
     file-path : Destination file path (parent directories are created as needed).
     data      : Vector of sheet maps. Each map must have :name and :sheet keys.
                 :sheet is a vector of maps (rows as maps) or vector of vectors
                 (first vector = headers, rest = data rows).
                 Add :table true (or :table {:name \"MyTable\" :style \"TableStyleMedium2\"})
                 to any sheet to render its data as a named Excel table."
  [file-path data]
  (let [_ (io/make-parents file-path)
        ;; Assign sequential table IDs to sheets that declare a table
        table-counter (atom 0)
        data-indexed  (mapv (fn [sheet]
                              (if (:table sheet)
                                (assoc sheet :_tid (swap! table-counter inc))
                                sheet))
                            data)
        table-ids     (keep :_tid data-indexed)
        workbook-xml  (str xmlh (hc/html [:workbook xlns
                                          (into [:sheets]
                                                (map-indexed #(vector :sheet
                                                                       {:name     (:name %2)
                                                                        :sheetId  (inc %)
                                                                        :r:id     (str "rId" (inc %))})
                                                             data-indexed))]))]
    (with-open [fos (FileOutputStream. ^String file-path)
                zos (ZipOutputStream. fos)]
      ;; Write worksheet XML for each sheet
      (dorun (map-indexed
              (fn [i sheet]
                (create-zip-entry zos (str "xl/worksheets/sheet" (inc i) ".xml")
                                  (if (:_tid sheet)
                                    (create-table-sheet-xml sheet)
                                    (create-sheet-xml sheet))))
              data-indexed))
      ;; Write table XML and worksheet relationship files for table sheets
      (dorun (map-indexed
              (fn [i sheet]
                (when-let [tid (:_tid sheet)]
                  (let [table-val  (:table sheet)
                        sheet-rows (:sheet sheet)
                        headers    (extract-table-headers sheet-rows)
                        row-count  (table-row-count sheet-rows)
                        tname      (cond
                                     (string? table-val) (sanitize-table-name table-val)
                                     (map? table-val)    (sanitize-table-name
                                                          (get table-val :name (str "Table" tid)))
                                     :else               (str "Table" tid))
                        tstyle     (if (map? table-val)
                                     (get table-val :style "TableStyleMedium2")
                                     "TableStyleMedium2")]
                    (create-zip-entry zos (str "xl/tables/table" tid ".xml")
                                      (make-table-xml tid tname tstyle headers row-count))
                    (create-zip-entry zos (str "xl/worksheets/_rels/sheet" (inc i) ".xml.rels")
                                      (make-sheet-rels-xml tid)))))
              data-indexed))
      ;; Write package metadata
      (create-zip-entry zos "[Content_Types].xml" (content-types (count data-indexed) table-ids))
      (create-zip-entry zos "_rels/.rels" wb-relationships)
      (create-zip-entry zos "xl/_rels/workbook.xml.rels" (ws-relationships (count data-indexed)))
      (create-zip-entry zos "xl/workbook.xml" workbook-xml))))