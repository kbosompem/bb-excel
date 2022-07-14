(ns bb-excel.core
  (:require [clojure.data.xml  :refer [parse-str]]
            [clojure.set       :refer [rename-keys]]
            [clojure.java.io   :refer [file]]
            [clojure.string    :refer [trim]]
            [clojure.pprint :refer [pprint]])
  (:import [java.time LocalDate Month]
           [java.time.format DateTimeFormatter]
           [java.util.zip  ZipFile])
  (:gen-class))

(set! *warn-on-reflection* true)

(def defaults
  "Default values for processing the Excel Spreadsheet
   :row  integer  :-  Which row to begin data extraction defaults to 1 
   :fxn  function :-  Which function to use parse header rows
   :rows integer  :-  Number of rows to extract
   :hdr  boolean  :- Rename columns with data from the first row"
  {:row 0
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
      (and (number? m) (pos? m)) (.plusDays (LocalDate/of 1899 Month/DECEMBER 30) m)
      :else n)))

(def error-codes
  {"#NAME?"   :bad-name
   "#DIV/0!"  :div-by-0
   "#REF!"    :invalid-reference
   "#NUM!"    :infinity
   "#N/A"     :not-applicable
   "#VALUE!"  :invalid-value
   "#NULL!"   :null
   nil :unknown-error})

(def date-formats {"14"  "dd/MM/yyyy"
                   "15"  "d-MMM-yy"
                   "16"  "d-MMM"
                   "17"  "MMM-yy"
                   "18"  "h:mm AM/PM"
                   "19"  "h:mm:ss AM/PM"
                   "20"  "h:mm"
                   "21"  "h:mm:ss"
                   "22"  "M/d/yy h:mm"
                   "30"  "M/d/yy"
                   "34"  "yyyy-MM-dd"
                   "45"  "mm:ss"
                   "46"  "[h]:mm:ss"
                   "47"  "mmss.0"
                   "51"  "MM-dd"
                   "52"  "yyyy-MM-dd"
                   "53"  "yyyy-MM-dd"
                   "55"  "yyyy-MM-dd"
                   "56"  "yyyy-MM-dd"
                   "58"  "MM-dd"
                   "165"  "M/d/yy"
                   "166"  "dd MMMM yyyy"
                   "167"  "dd/MM/yyyy"
                   "168"  "dd/MM/yy"
                   "169"  "d.M.yy"
                   "170"  "yyyy-MM-dd"
                   "171"  "dd MMMM yyyy"
                   "172"  "d MMMM yyyy"
                   "173"  "M/d"
                   "174"  "M/d/yy"
                   "175"  "MM/dd/yy"
                   "176"  "d-MMM"
                   "177"  "d-MMM-yy"
                   "178"  "dd-MMM-yy"
                   "179"  "MMM-yy"
                   "180"  "MMMM-yy"
                   "181"  "MMMM d, yyyy"
                   "182"  "M/d/yy hh:mm t"
                   "183"  "M/d/y HH:mm"
                   "184"  "MMM"
                   "185"  "MMM-dd"
                   "186"  "M/d/yyyy"
                   "187" "d-MMM-yyyy"})

(defn bbe-format
  [styles nfs {:keys [s d] :as c}]
  (println :c c)
  (let [df (styles (max 0 (dec (parse-long s))))
        style (nfs df)]
    (cond
      style (try (.format (num2date d) (DateTimeFormatter/ofPattern style))
                 (catch Exception _ (println "Unable to format" c df style) d))
      :else d)))

(defn process-cell
  "Process Excel cell"
  [dict styles nfs coll]
  (clojure.pprint/pprint coll)
  (let [[[_ row col]] (re-seq #"([A-Z]*)([0-9]+)" (:r coll))
        u (-> coll
              (assoc :x row)
              (assoc :y col))]
    (cond
      (= (:t u) "s")      (dissoc (assoc-in u [:d] (dict (read-string (:d u)))) :t)
      (= (:t u) "str")    (dissoc u :t)
      (= (:t u) "b")      (dissoc (assoc-in u [:d] (if (= "1" (:d u)) true false)) :t)
      (= (:t u) "e")      (assoc-in u [:d] (error-codes (:d u)))
      (not (nil? (:s u))) (assoc-in u [:d] (bbe-format styles nfs u))
      :else (update-in u [:d] s2d))))

(defn process-row
  "Process Excel row of data"
  [dict styles nfs coll]
  (reduce #(merge % {:_r (read-string (:y %2)) (keyword (:x %2)) (:d %2)}) {}
          (map (partial process-cell dict styles nfs)
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
  [filename]
  (let [zf (ZipFile. ^String filename)
        wb (.getEntry zf (str "xl/sharedStrings.xml"))
        ins (.getInputStream zf wb)
        x (parse-str (slurp ins))]
    (->>
     (filter #((:text-part tags) (:tag %)) (xml-seq x))
     (map get-cell-text)
     ;(map (comp last :content))
     (zipmap (range)))))

(defn get-num-format-map
  "Get styles"
  [filename]
  (let [zf  (ZipFile. ^String filename)
        wb  (.getEntry zf (str "xl/styles.xml"))
        ins (.getInputStream zf wb)
        x   (parse-str (slurp ins))
        s   (->> x
                 xml-seq
                 (filter #((:cellxfs tags) (:tag %)))
                 first
                 :content
                 (filter #((:xf tags) (:tag %)))
                 (mapv (comp :numFmtId :attrs)))
        _ (pprint s)
        nfs (merge (->> x
                        xml-seq
                        (filter #((:numFmts tags) (:tag %)))
                        first
                        :content
                        (map :attrs)
                        (map #(hash-map (:numFmtId %) (:formatCode %)))
                        (apply merge))
                   date-formats)
        _ (pprint nfs)]
    [s  nfs]))

(defn get-sheet
  "Get sheet from file"
  ([filename sheetname]
   (get-sheet filename sheetname {}))
  ([filename sheetname options]
   (let [opts    (merge defaults options)
         row     (:row opts)
         hdr     (:hdr opts)
         row     (if (and hdr (zero? row)) 1 row)
         rows    (:rows opts)
         fxn     (:fxn opts)
         cols    (map fxn (:columns opts))
         sheetid (:idx (first (filter #(= sheetname (:name %)) (get-sheet-names filename))))
         zf      (ZipFile. ^String filename)
         wb      (.getEntry zf (str "xl/worksheets/sheet" sheetid ".xml"))
         ins     (.getInputStream zf wb)
         dict    (get-unique-strings filename)
         [styles nfs]  (get-num-format-map filename)
         xx      (slurp ins)
         x       (parse-str xx)
         d       (->>  x :content
                       (filter #((:sheet-data tags) (:tag %)))
                       first :content
                       (map :content)
                       (take rows)
                       (map (partial process-row dict styles nfs)))
         dx (remove #(= row (:_r %)) d)
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
  (let [[[_ sc sr ec er]] (re-seq #"([A-Z]+)([1-9]*)[:]?([A-Z]*)([1-9]*)" s)
        ec (or (when-str ec) sc)
        sr (or (when-num sr) 1)
        er (or (when-num er) 10000)]
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


(comment

  (->>
   (get-sheets "test/data/Types.xlsx" {:hdr false :rows 1000 :row 0})
   second
   :sheet
   clojure.pprint/print-table)

  #{})