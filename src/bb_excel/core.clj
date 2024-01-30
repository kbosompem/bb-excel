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

(defonce ^SimpleDateFormat sdf (SimpleDateFormat. "HH:mm:ss"))
(.setTimeZone sdf (TimeZone/getTimeZone "UTC"))

(defn parse-xlong [x]
  (parse-long (or x "")))

(def error-codes
  {"#NAME?"   :bad-name
   "#DIV/0!"  :div-by-0
   "#REF!"    :invalid-reference
   "#NUM!"    :infinity
   "#N/A"     :not-applicable
   "#VALUE!"  :invalid-value
   "#NULL!"   :null
   "#SPILL!"  :multiple-results
   nil        :unknown-error})

(def defaults
  "Default values for processing the Excel Spreadsheet
   :row  integer  :-  Which row to begin data extraction defaults to 0 
   :fxn  function :-  Which function to use parse header rows
   :rows integer  :-  Number of rows to extract
   :hdr  boolean  :- Rename columns with data from the first row"
  {:row 0
   :fxn str
   :rows 10000
   :hdr false})

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
                       (map #(select-keys % [:sheetId :name]))
                       (map #(update % :sheetId parse-xlong))
                       (map #(rename-keys % {:sheetId :idx})))
              sheet-nodes)))
    []))

(defn get-sheet-names
  "Retrieves a list of Sheet Names from a given Excel Spreadsheet"
  [file-or-filename]
  (let [^ZipFile zipfile (get-zipfile file-or-filename)]
    (get-sheet-names* zipfile)))

(defn num2date
  "Format Excel Date"
  [n]
  (when n (.format (.plusDays (LocalDate/of 1899 Month/DECEMBER 30)  (parse-double (str n)))
                   (DateTimeFormatter/ofPattern "MM/dd/yyyy"))))

(defn num2time
  "Format Excel Time"
  [n]
  (when n (.format sdf (*  (parse-double (str n)) 24 60 60 1000))))

(defn num2pct
  "Format Percentage"
  [n]
  (when n (format "%.4f%%" (* 100 (parse-double (str n))))))

(def dates #{"14"  "15"  "16"  "17"  "30"  "34"  "51"
             "52"  "53"  "55"  "56"  "58"  "165"
             "166" "167" "168" "169" "170" "171" "172"
             "173" "174" "175" "176" "177" "178" "179"
             "180" "181" "184" "185" "186" "187"})

(def times #{"164"  "18" "19" "21" "20"  "45" "46" "47"})

(def pcts  #{"9" "10"})

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

(def ^:const BASE_ROW_INDEX 0)
(def ^:const BASE_COLUMN_INDEX 0)

(defn valid-cell-index?
  [cell-index]
  (if cell-index
    (boolean (re-find #"^[A-Z]{1,3}\d+$" cell-index))
    false))

(def ^:const A_CHAR_INDEX (int \A))

(defn number->column-letter
  [n]
  (loop [num n
         acc ""]
    (if (> num 0)
      (let [residue (mod (dec num) 26)
            new-num (quot (dec num) 26)]
        (recur new-num (str (char (+ residue A_CHAR_INDEX)) acc)))
      acc)))

(defn get-col-index
  "Self-calculated index is used only if cell-index attribute(:r) is missing on the cell"
  [cell last-processed-col-index]
  (let [cell-index (-> cell :attrs :r)]
    (if (valid-cell-index? cell-index)
      (re-find #"[A-Z]{1,3}" cell-index) 
          (-> last-processed-col-index
          (inc)
          (number->column-letter)))))

(defn process-row
  "Process Excel row of data"
  [shared-strings styles row]
  (->> (:content row)
       (reduce (fn [{:keys [row-data last-processed-col-index]} cell]
                 (let [col-index (get-col-index cell last-processed-col-index)
                       cell-value (extract-cell-value shared-strings styles cell)]
                   {:row-data (assoc row-data (keyword col-index) cell-value)
                    :last-processed-col-index col-index}))
               {:row-data {}
                :last-processed-col-index BASE_COLUMN_INDEX})
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

(defn get-and-check-sheet-id
  [^ZipFile zipfile sheetname-or-idx]
  (let [sheets (get-sheet-names* zipfile)
        found-sheet
        (find-first (fn [sheet]
                      (cond
                        (string? sheetname-or-idx)
                        (= (str/lower-case sheetname-or-idx)
                           (str/lower-case (:name sheet)))

                        (and (integer? sheetname-or-idx)
                             (pos? sheetname-or-idx))
                        (= sheetname-or-idx (:idx sheet))))
                    sheets)]
    (or (:idx found-sheet)
        (throw-ex (format "Could not find sheet with name or index equal '%s'! Sheet does not exist." sheetname-or-idx)))))

(defn get-sheet-entry
  [^ZipFile zipfile ^long sheet-id]
  (or (.getEntry zipfile (str "xl/worksheets/sheet" sheet-id ".xml"))
      (throw-ex (format "Could not find sheet with sheet-id equal '%s'! Sheet data file does not exist." sheet-id))))

(defn get-sheet
  "Get sheet from file or filename"
  ([file-or-filename]
   (get-sheet file-or-filename 1 {}))
  ([file-or-filename sheetname-or-idx]
   (get-sheet file-or-filename sheetname-or-idx {}))
  ([file-or-filename sheetname-or-idx options]
   (let [^ZipFile zipfile (get-zipfile file-or-filename)
         ^long sheet-id (get-and-check-sheet-id zipfile sheetname-or-idx)
         ^ZipEntry sheet-entry (get-sheet-entry zipfile sheet-id)
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

; -------------------------------------
; EXPERIMENTAL 
; --------------------------------------
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

(defn ws-relationships [n]
  (str xmlh
       (hc/html
        (into [:Relationships {:xmlns "http://schemas.openxmlformats.org/package/2006/relationships"}]
              (for [x (range n)]
                [:Relationship {:Id (str "rId" (inc x))
                                :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                                :Target (str "worksheets/sheet" (inc x) ".xml")}])))))

(defn content-types [n]
  (str xmlh
       (hc/html
        (into [:Types {:xmlns "http://schemas.openxmlformats.org/package/2006/content-types"}
               [:Default {:Extension :rels
                          :ContentType "application/vnd.openxmlformats-package.relationships+xml"}]
               [:Default {:Extension :xml
                          :ContentType :application/xml}]
               [:Override {:PartName "/xl/workbook.xml"
                           :ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"}]]
              (for [x (range n)]
                [:Override {:PartName (str "/xl/worksheets/sheet" (inc x) ".xml")
                            :ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"}])))))

(defn excel-date-serial [datetime]
  (.between ChronoUnit/DAYS (LocalDate/of 1899 Month/DECEMBER 30) datetime))

(defn excel-time-serial [datetime]
  (/ (.between ChronoUnit/SECONDS (LocalDateTime/of 1899 Month/DECEMBER 30 0 0) datetime) 86400.0))

(defn cell-type [value]
  (cond
    (instance? java.time.LocalDate value) ["n" (excel-date-serial value)]
    (instance? java.time.LocalDateTime value) ["n" (excel-time-serial value)]
    (string? value) ["inlineStr" [:is [:t value]]]
    (number? value) ["n" value]
    (boolean? value) ["b" (if value "1" "0")]
    :else ["inlineStr" [:is [:t (str value)]]]))

(defn generate-xml-cell
  [col-letter row-num value]
  (let [[cell-type cell-value] (cell-type value)]
    [:c {:r (str (if (keyword? col-letter) (name col-letter) col-letter)
                   (inc row-num)) :t cell-type} cell-value]))


(defn generate-xml-row
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

(defn create-sheet-xml 
  [data]
  (let [headers (if (map? data)
                  (keys (first (:sheet data)))
                  (first data))
        rows (if (map? data)
               (map-indexed #(generate-xml-row %2 (inc %) (:cmap data) ) (:sheet data))
               (map (comp (partial generate-xml-row) vals) (:sheet data)))]
    (str (hc/html [:worksheet xlns
                   [:sheetData (cons (generate-xml-row headers 0) rows)]]))))

(defn create-zip-entry 
  [zip-stream entry-name content]
  (let [entry  (ZipEntry. ^String entry-name)]
    (.putNextEntry ^ZipOutputStream zip-stream ^ZipEntry entry)
    (.write ^ZipOutputStream zip-stream (.getBytes ^String content "UTF-8"))
    (.closeEntry ^ZipOutputStream zip-stream)))

(defn create-xlsx 
  [file-path data]
  (let [workbook-xml  (str xmlh (hc/html [:workbook xlns
                                          (into [:sheets]
                                                (map-indexed #(vector :sheet {:name (:name %2) :sheetId (inc %)
                                                                              :r:id (str "rId" (inc %))}) data))]))]
    (with-open [fos (FileOutputStream. ^String file-path)
                zos (ZipOutputStream. fos)]
      (dorun (map-indexed
              #(create-zip-entry zos (str "xl/worksheets/sheet" (inc %) ".xml")  (create-sheet-xml %2)) data))
      (create-zip-entry zos "[Content_Types].xml" (content-types (count data)))
      (create-zip-entry zos "_rels/.rels" wb-relationships)
      (create-zip-entry zos "xl/_rels/workbook.xml.rels" (ws-relationships (count data)))
      (create-zip-entry zos "xl/workbook.xml" workbook-xml))))

