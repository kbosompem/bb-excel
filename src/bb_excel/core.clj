(ns bb-excel.core
  (:require [clojure.data.xml  :refer [parse-str]]
            [clojure.java.io   :as io]
            [clojure.set       :refer [rename-keys]])
  (:import [java.io File]
           [java.text SimpleDateFormat]
           [java.time LocalDate Month]
           [java.time.format DateTimeFormatter]
           [java.util TimeZone]
           (java.util.zip ZipFile))
  (:gen-class))

(set! *warn-on-reflection* true)

(defonce ^SimpleDateFormat sdf (SimpleDateFormat. "HH:mm:ss"))
(.setTimeZone sdf (TimeZone/getTimeZone "UTC"))

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
  {:row 0 :fxn str :rows 10000 :hdr false})

(def tags
  "Map of Excel XML namespaces of interest"
  {:sheet-tag  #{:sheets
                 :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/sheets
                 :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/sheets}
   :row-tag    #{:row
                 :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/row
                 :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/row}
   :text-part   #{:si
                  :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/si
                  :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/si}
   :numFmts   #{:numFmts
                :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/numFmts
                :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/numFmts}
   :cellxfs   #{:cellXfs
                :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/cellXfs
                :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/cellXfs}
   :xf   #{:xf
           :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/xf
           :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/xf}
   :text-t   #{:t
               :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/t
               :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/t}
   :text-r   #{:r
               :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/r
               :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/r}
   :sheet-data #{:sheetData
                 :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/sheetData
                 :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/sheetData}
   :sheet-id   #{:id
                 :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2FofficeDocument%2F2006%2Frelationships/id
                 :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2FofficeDocument%2Frelationships/id}})

(defn- zipfile-or-nil
  "Retrieve ZipFile object if provided `file-or-filename` point to existing file or nil"
  [file-or-filename]
  (when-let [^File file (condp instance? file-or-filename
                          String (io/file file-or-filename)
                          File file-or-filename
                          nil)]
    (when (.exists file)
      (ZipFile. file))))

(defn get-sheet-names
  "Retrieves a list of Sheet Names from a given Excel Spreadsheet
   Returns nil if the file does not exist or a non-string is passed as the `file-or-filename`"
  [file-or-filename]
  (when-let [^ZipFile zipfile (zipfile-or-nil file-or-filename)]
    (let [wb (.getEntry zipfile "xl/workbook.xml")
          ins (.getInputStream zipfile wb)
          x (parse-str (slurp ins))
          y (filter #((:sheet-tag tags) (:tag %)) (xml-seq x))]
      (->> y
           first
           :content
           (map :attrs)
           (map-indexed #(select-keys
                          (rename-keys
                           (assoc %2 :idx (inc %))
                           (zipmap (:sheet-id tags) (repeat :id)))
                          [:id :name :sheetId :idx]))))))

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
  [coll styles ids]
  (when (:s coll)
    (try
      (ids (styles (parse-long (:s coll))))
      (catch Exception _ false))))

(defn process-cell
  "Process Excel cell"
  [dict styles coll]
  (let [[[_ row col]] (re-seq #"([A-Z]*)([0-9]+)" (:r coll))
        u (-> coll
              (assoc :x row)
              (assoc :y col))]
    (cond
      ;; Possible data types well explained here https://stackoverflow.com/a/18346273
      (= (:t u) "s")                (dissoc (assoc-in u [:d] (dict (read-string (:d u)))) :t)
      (= (:t u) "str")              (dissoc u :t)
      (= (:t u) "inlineStr")        (dissoc u :t)
      (= (:t u) "b")                (dissoc (assoc-in u [:d] (if (= "1" (:d u)) true false)) :t)
      (= (:t u) "e")                (assoc-in u [:d] (error-codes (:d u)))
      (= (:t u) "n")                (assoc u :d (parse-long (:d u)))
      (style-check u styles pcts)   (assoc-in u [:d] (num2pct (:d u)))
      (style-check u styles dates)  (assoc-in u [:d] (num2date (:d u)))
      (style-check u styles times)  (assoc-in u [:d] (num2time (:d u)))
      :else u)))

(defn process-row
  "Process Excel row of data"
  [dict styles coll]
  (reduce #(merge % {:_r (read-string (:y %2)) (keyword (:x %2)) (:d %2)}) {}
          (map (partial process-cell dict styles)
               (map #(merge (first %) {:d (second %)}
                            {:f (nth % 2)})
                    (map (juxt :attrs
                               (comp last :content last :content)
                               (comp first :content first :content)) coll)))))

(defn- get-cell-text
  "Extract "
  [coll]
  (apply str
         (mapcat :content
                 (filter #((:text-t tags) (:tag %))
                         (xml-seq coll)))))

(defn get-unique-strings
  "Get dictionary of all unique strings in the Excel spreadsheet"
  [^ZipFile zipfile]
  (if-let [wb (.getEntry zipfile (str "xl/sharedStrings.xml"))]
    (let [ins (.getInputStream zipfile wb)
          x (parse-str (slurp ins))]
      (->>
       (filter #((:text-part tags) (:tag %)) (xml-seq x))
       (map get-cell-text)
       (zipmap (range))))
    {}))

(defn get-styles
  "Get styles"
  [^ZipFile zipfile]
  (let [wb  (.getEntry zipfile (str "xl/styles.xml"))
        ins (.getInputStream zipfile wb)
        x   (parse-str (slurp ins))]
    (->> x
         xml-seq
         (filter #((:cellxfs tags) (:tag %)))
         first
         :content
         (filter #((:xf tags) (:tag %)))
         (mapv (comp :numFmtId :attrs)))))

(defn get-sheet
  "Get sheet from file or filename"
  ([file-or-filename]
   (get-sheet file-or-filename 1 {}))
  ([file-or-filename sheetname-or-idx]
   (get-sheet file-or-filename sheetname-or-idx {}))
  ([file-or-filename sheetname-or-idx options]
   (if-let [^ZipFile zipfile (zipfile-or-nil file-or-filename)]
     (let [opts    (merge defaults options)
           row     (:row opts)
           hdr     (:hdr opts)
           row     (if (and hdr (zero? row)) 1 row)
           rows    (:rows opts)
           fxn     (:fxn opts)
           cols    (map fxn (:columns opts))
           sheetid (cond
                     (string? sheetname-or-idx)
                     (:idx (first (filter #(= sheetname-or-idx (:name %)) (get-sheet-names file-or-filename))))

                     (and (integer? sheetname-or-idx) (pos? sheetname-or-idx))
                     sheetname-or-idx

                     :else
                     (let [message (format "Attr 'sheetname-or-idx' can only be string or positive number, but passed '%s'" sheetname-or-idx)]
                       (throw (ex-info message {}))))
           wb      (.getEntry zipfile (str "xl/worksheets/sheet" sheetid ".xml"))
           ins     (.getInputStream zipfile wb)
           dict    (get-unique-strings zipfile)
           styles  (get-styles zipfile)
           xx      (slurp ins)
           x       (parse-str xx)
           d       (->>  x :content
                         (filter #((:sheet-data tags) (:tag %)))
                         first :content
                         (map :content)
                         (take rows)
                         (map (partial process-row dict styles)))
           dx (remove #(= row (:_r %)) d)
           h (when hdr (merge (update-vals (first (filter #(= (:_r %) row) d)) fxn) {:_r :_r}))
           dy (if (pos? rows)
                (take rows (map #(rename-keys % h) dx))
                (map #(rename-keys % h) dx))]
       (if (empty? cols) dy (map #(select-keys % cols) dy)))
     (let [message (format "Attr 'file-or-filename' contains value not suitable for creating ZipFile: '%s'" file-or-filename)]
       (throw (ex-info message {}))))))


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
  (let [[[_ osc osr oec oer]] (re-seq #"([A-Z]+)([0-9]*)[:]?([A-Z]*)([0-9]*)" s)
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
       (filter #((set rows) (:_r %)) sheet)))

(defn get-range
  "Get range of values using Excel cell coordinates
   e.g A1:C5"
  [sheet rg]
  (let [{:keys [cols rows]} (parse-range rg)
        [rs re] rows
        [cs ce] cols]
    (get-cells sheet (range rs re) (crange cs ce))))
