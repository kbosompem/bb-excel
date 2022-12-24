(ns bb-excel.cli
  (:require [bb-excel.core     :refer [get-sheets]]
            [clojure.java.io   :refer [make-parents]]
            [clojure.tools.cli :refer [parse-opts]]
            [clojure.string    :refer [split join lower-case trim] :as str]
            [clojure.pprint    :refer [print-table]])
  (:gen-class))

(set! *warn-on-reflection* true)

(defn left
  "Takes up to n characters from a string"
  [s n]
  (cond (and s (string? s))
        (subs s 0 (max 0 (min (count s) n)))
        (coll? s) (take n s)))

(defn skeyword
  "Sanitizes column headers and converts them to keywords.
   1. Removes slashes and spaces
   2. Lower cases 
   3. Takes up to 50 characters
   4. Replaces non-ascii characters with an underscore.
   
   **Not appropriate for foreign language headers"
  [s]
  (keyword
   (left
    (str/replace
     (str/replace (trim (lower-case (str s))) #"[\[\]]" "")
     #"[^A-Za-z0-9\-]+" "_") 50)))

(def fxns
  "Map of functions"
  {:str str
   "str" str
   :keyword keyword
   "keyword" keyword
   :skeyword skeyword
   "skeyword" skeyword
   nil str})

(defn bbexcel
  "Extract Excel Sheets into EDN"
  [input output options]
  (when-not
   (or (nil? input) (nil? output))
    (make-parents output)
    (spit output (pr-str (get-sheets input options)))))

(def cli-options
  "Command Line Options"
  [["-d" "--hdr" "Use Header Row"]
   ["-r" "--row r" "Start Row"
    :parse-fn #(parse-long %)
    :desc "Start Row"]
   ["-n" "--rows s" "End Row"
    :parse-fn #(parse-long %)
    :desc "End Row"]
   ["-f" "--fxn f" "Function"
    :parse-fn fxns
    :desc "Parser"]
   ["-s" "--sheet x" "Sheet"
    :desc "Sheet Name"]
   ["-p" "--print"
    :desc "Print Tables"]
   ["-c" "--columns c"
    :parse-fn #(map keyword (split % #" "))
    :desc "Columns"]
   ["-h" "--help"]])

(defn error-msg
  "Error messages"
  [errors]
  (str "The following errors occurred while parsing your command:\n"
       (join \newline errors)))

(defn help
  "Command line options"
  [summary]
  (->> ["bbexcel"
        ""
        "Usage: bb input-file output-file options"
        ""
        "Options:"
        summary
        ""
        "Please refer to the manual page for more information."]
       (join \newline)))

(defn -main [& args]
  (let [{:keys [options arguments summary errors]}
        (parse-opts args cli-options)
        [input output] arguments]
    (cond
      errors (println (error-msg errors))
      (and (empty? options) (nil? output)) (println (help summary))
      (or (nil? (first args))
          (:help options)) (println (help summary))
      (:print options) (doseq [y (get-sheets input options)]
                         (println :SHEET= (:name y))
                         (print-table (:sheet y)))
      :else (bbexcel input output options))))

