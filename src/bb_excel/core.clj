(ns bb-excel.core
  (:require [clojure.data.xml  :refer [parse-str]]
            [clojure.set       :refer [rename-keys]]
            [clojure.java.io   :refer [file]]
            [clojure.string    :refer [trim]])
  (:import [java.time LocalDate Month]
           [java.util.zip  ZipFile])
  (:gen-class))

(set! *warn-on-reflection* true)

(def defaults
  "Default values for processing the Excel Spreadsheet
   :row  integer  :-  Which row to begin data extraction defaults to 1 
   :fxn  function :-  Which function to use parse header rows
   :rows integer  :-  Number of rows to extract
   :hdr  boolean  :- Rename columns with data from the first row"
  {:row 1
   :fxn str
   :hdr false
   :rows 10000})

(def tags
  "Map of Excel XML namespaces of interest"
  {:sheet-tag  #{:sheets
                 :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/sheets
                 :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/sheets}
   :row-tag    #{:row
                 :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/row
                 :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/row}
   :text-tag   #{:t
                 :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/t
                 :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/t}
   :sheet-data #{:sheetData
                 :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2Fspreadsheetml%2F2006%2Fmain/sheetData
                 :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2Fspreadsheetml%2Fmain/sheetData}
   :sheet-id   #{:id
                 :xmlns.http%3A%2F%2Fschemas.openxmlformats.org%2FofficeDocument%2F2006%2Frelationships/id
                 :xmlns.http%3A%2F%2Fpurl.oclc.org%2Fooxml%2FofficeDocument%2Frelationships/id}})


(defn get-sheet-names
  "Retrieves a list of Sheet Names from a given Excel Spreadsheet
   Returns nil if the file does not exist or a non-string is passed as the filename"
  [filename]
  (when (and (not-any? (fn [f] (f filename)) [nil? coll?])
             (.exists (file filename)))
  (let [^ZipFile zf (ZipFile. ^String filename)
        wb (.getEntry zf "xl/workbook.xml")
        ins (.getInputStream zf wb)
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

(defn s2d
  "Convert string to double"
  [n]
  (if (string? n)
    (Double/parseDouble (trim n))
    n))

(defn num2date
  "Convert number to Excel compatible Date"
  [n]
  (let [m (cond
            (string? n) (parse-double n)
            :else n)]
    (cond
      (and (number? m) (pos? m)) (.toString (.plusDays (LocalDate/of 1899 Month/DECEMBER 30) m))
      :else n)))

(defn process-cell
  "Process Excel cell"
  [dict coll]
  (let [[[_ row col]] (re-seq #"([A-Z]*)([0-9]+)" (:r coll))
        u (-> coll
              (assoc :x row)
              (assoc :y col))]
    (cond
      (= (:t u) "s")    (dissoc (assoc-in u [:d] (dict (read-string (:d u)))) :t)
      (= (:t u) "str")  u
      (= (:t u) "e")    (assoc-in u [:d] (str "Error : " (:d u)))
      (= (:s u) "1")    (assoc-in u [:d] (num2date (:d u)))
      :else (update-in u [:d] s2d))))

(defn process-row
  "Process Excel row of data"
  [dict coll]
  (reduce #(merge % {:_r (read-string (:y %2)) (keyword (:x %2)) (:d %2)}) {}
          (map (partial process-cell dict)
               (map #(merge (first %) {:d (second %)}
                            {:f (nth % 2)})
                    (map (juxt :attrs
                               (comp last :content last :content)
                               (comp first :content first :content)) coll)))))

(defn get-unique-strings
  "Get dictionary of all unique strings in the Excel spreadsheet"
  [filename]
  (let [zf (ZipFile. ^String filename)
        wb (.getEntry zf (str "xl/sharedStrings.xml"))
        ins (.getInputStream zf wb)
        x (parse-str (slurp ins))]
    (->>
     (filter #((:text-tag tags) (:tag %)) (xml-seq x))
     (map (comp first :content))
     (zipmap (range)))))


(defn get-sheet
  "Get sheet from file"
  ([filename sheetname]
   (get-sheet filename sheetname {}))
  ([filename sheetname options]
   (let [opts    (merge defaults options)
         row     (:row opts)
         hdr     (:hdr opts)
         rows    (:rows opts)
         fxn     (:fxn opts)
         cols    (map fxn (:columns opts))
         sheetid (:idx (first (filter #(= sheetname (:name %)) (get-sheet-names filename))))
         zf      (ZipFile. ^String filename)
         wb      (.getEntry zf (str "xl/worksheets/sheet" sheetid ".xml"))
         ins     (.getInputStream zf wb)
         dict    (get-unique-strings filename)
         xx      (slurp ins)
         x       (parse-str xx)
         d       (->>  x :content
                       (filter #((:sheet-data tags) (:tag %)))
                       first :content
                       (map :content)
                       (take rows)
                       (map (partial process-row dict)))
         dx (remove #(= 0 (:_r %)) d)
         h (when hdr (merge (update-vals (first (filter #(= (:_r %) row) d)) fxn) {:_r :_r}))
         dy (if (pos? rows)
              (take rows (map #(rename-keys % h) dx))
              (map #(rename-keys % h) dx))]
     (if (empty? cols) dy (map #(select-keys % cols) dy)))))

(defn get-sheets
  "Get all or specified sheet from the excel spreadsheet"
  ([filename]
   (get-sheets filename {}))
  ([filename options]
   (let [sns  (get-sheet-names filename)
         sxs  (if (:sheet options) (filter #(= (:sheet options) (:name %)) sns) sns)
         res  (if (empty? sxs) [{:sheet []}]
                  (map #(assoc % :sheet
                               (try (get-sheet filename (:name %) options)
                                    (catch Exception ex [(bean ex)]))) sxs))]
     res)))

(defn get-row
  "Get row from sheet by row index"
  [sheet row]
  (first (filter #(= row (:_r %)) sheet)))

(defn get-col
  "Get column from sheet by name. 
   If columns have been renamed use the new name."
  [sheet col]
  (map #(select-keys % [:_r col]) sheet))

(defn get-range
  "Get range of values returned as list of rows"
  [sheet rows cols]
  (map #(select-keys % cols)
       (filter #((set rows) (:_r %))  sheet)))

