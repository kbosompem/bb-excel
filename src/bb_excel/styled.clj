(ns bb-excel.styled
  "Experimental: Create styled Excel spreadsheets using Tailwind-like CSS classes.

   Syntax:
     [:A1.bg-blue-500.text-white.font-bold \"Header\"]     ; Single cell
     [:Sheet1/A1.bg-red-500 \"Value\"]                     ; With sheet name
     [:5.bg-gray-100 {:A \"Col1\" :B \"Col2\"}]            ; Entire row
     [:AA.border-right {:1 \"R1\" :2 \"R2\"}]              ; Entire column
     [:A1:C3.bg-yellow-200 [[1 2 3] [4 5 6] [7 8 9]]]     ; Range

   Supported Tailwind classes (subset):
     Background: bg-{color}-{shade}
     Text color: text-{color}-{shade}
     Text size:  text-xs, text-sm, text-base, text-lg, text-xl, text-2xl
     Font:       font-bold, font-normal, italic, underline
     Alignment:  text-left, text-center, text-right
     Borders:    border, border-{side}, border-{color}-{shade}
     Border style: border-dashed, border-dotted, border-double"
  (:require [clojure.string :as str]
            [hiccup2.core :as hc]
            [clojure.java.io :as io]
            [bb-excel.core :as core])
  (:import [java.io FileOutputStream]
           [java.util.zip ZipEntry ZipOutputStream]))

;; Tailwind color palette (500 shades as default, with variants)
(def ^:private colors
  {"slate"   {:50 "F8FAFC" :100 "F1F5F9" :200 "E2E8F0" :300 "CBD5E1" :400 "94A3B8"
              :500 "64748B" :600 "475569" :700 "334155" :800 "1E293B" :900 "0F172A"}
   "gray"    {:50 "F9FAFB" :100 "F3F4F6" :200 "E5E7EB" :300 "D1D5DB" :400 "9CA3AF"
              :500 "6B7280" :600 "4B5563" :700 "374151" :800 "1F2937" :900 "111827"}
   "zinc"    {:50 "FAFAFA" :100 "F4F4F5" :200 "E4E4E7" :300 "D4D4D8" :400 "A1A1AA"
              :500 "71717A" :600 "52525B" :700 "3F3F46" :800 "27272A" :900 "18181B"}
   "neutral" {:50 "FAFAFA" :100 "F5F5F5" :200 "E5E5E5" :300 "D4D4D4" :400 "A3A3A3"
              :500 "737373" :600 "525252" :700 "404040" :800 "262626" :900 "171717"}
   "stone"   {:50 "FAFAF9" :100 "F5F5F4" :200 "E7E5E4" :300 "D6D3D1" :400 "A8A29E"
              :500 "78716C" :600 "57534E" :700 "44403C" :800 "292524" :900 "1C1917"}
   "red"     {:50 "FEF2F2" :100 "FEE2E2" :200 "FECACA" :300 "FCA5A5" :400 "F87171"
              :500 "EF4444" :600 "DC2626" :700 "B91C1C" :800 "991B1B" :900 "7F1D1D"}
   "orange"  {:50 "FFF7ED" :100 "FFEDD5" :200 "FED7AA" :300 "FDBA74" :400 "FB923C"
              :500 "F97316" :600 "EA580C" :700 "C2410C" :800 "9A3412" :900 "7C2D12"}
   "amber"   {:50 "FFFBEB" :100 "FEF3C7" :200 "FDE68A" :300 "FCD34D" :400 "FBBF24"
              :500 "F59E0B" :600 "D97706" :700 "B45309" :800 "92400E" :900 "78350F"}
   "yellow"  {:50 "FEFCE8" :100 "FEF9C3" :200 "FEF08A" :300 "FDE047" :400 "FACC15"
              :500 "EAB308" :600 "CA8A04" :700 "A16207" :800 "854D0E" :900 "713F12"}
   "lime"    {:50 "F7FEE7" :100 "ECFCCB" :200 "D9F99D" :300 "BEF264" :400 "A3E635"
              :500 "84CC16" :600 "65A30D" :700 "4D7C0F" :800 "3F6212" :900 "365314"}
   "green"   {:50 "F0FDF4" :100 "DCFCE7" :200 "BBF7D0" :300 "86EFAC" :400 "4ADE80"
              :500 "22C55E" :600 "16A34A" :700 "15803D" :800 "166534" :900 "14532D"}
   "emerald" {:50 "ECFDF5" :100 "D1FAE5" :200 "A7F3D0" :300 "6EE7B7" :400 "34D399"
              :500 "10B981" :600 "059669" :700 "047857" :800 "065F46" :900 "064E3B"}
   "teal"    {:50 "F0FDFA" :100 "CCFBF1" :200 "99F6E4" :300 "5EEAD4" :400 "2DD4BF"
              :500 "14B8A6" :600 "0D9488" :700 "0F766E" :800 "115E59" :900 "134E4A"}
   "cyan"    {:50 "ECFEFF" :100 "CFFAFE" :200 "A5F3FC" :300 "67E8F9" :400 "22D3EE"
              :500 "06B6D4" :600 "0891B2" :700 "0E7490" :800 "155E75" :900 "164E63"}
   "sky"     {:50 "F0F9FF" :100 "E0F2FE" :200 "BAE6FD" :300 "7DD3FC" :400 "38BDF8"
              :500 "0EA5E9" :600 "0284C7" :700 "0369A1" :800 "075985" :900 "0C4A6E"}
   "blue"    {:50 "EFF6FF" :100 "DBEAFE" :200 "BFDBFE" :300 "93C5FD" :400 "60A5FA"
              :500 "3B82F6" :600 "2563EB" :700 "1D4ED8" :800 "1E40AF" :900 "1E3A8A"}
   "indigo"  {:50 "EEF2FF" :100 "E0E7FF" :200 "C7D2FE" :300 "A5B4FC" :400 "818CF8"
              :500 "6366F1" :600 "4F46E5" :700 "4338CA" :800 "3730A3" :900 "312E81"}
   "violet"  {:50 "F5F3FF" :100 "EDE9FE" :200 "DDD6FE" :300 "C4B5FD" :400 "A78BFA"
              :500 "8B5CF6" :600 "7C3AED" :700 "6D28D9" :800 "5B21B6" :900 "4C1D95"}
   "purple"  {:50 "FAF5FF" :100 "F3E8FF" :200 "E9D5FF" :300 "D8B4FE" :400 "C084FC"
              :500 "A855F7" :600 "9333EA" :700 "7E22CE" :800 "6B21A8" :900 "581C87"}
   "fuchsia" {:50 "FDF4FF" :100 "FAE8FF" :200 "F5D0FE" :300 "F0ABFC" :400 "E879F9"
              :500 "D946EF" :600 "C026D3" :700 "A21CAF" :800 "86198F" :900 "701A75"}
   "pink"    {:50 "FDF2F8" :100 "FCE7F3" :200 "FBCFE8" :300 "F9A8D4" :400 "F472B6"
              :500 "EC4899" :600 "DB2777" :700 "BE185D" :800 "9D174D" :900 "831843"}
   "rose"    {:50 "FFF1F2" :100 "FFE4E6" :200 "FECDD3" :300 "FDA4AF" :400 "FB7185"
              :500 "F43F5E" :600 "E11D48" :700 "BE123C" :800 "9F1239" :900 "881337"}
   ;; Special colors
   "white"   {:500 "FFFFFF"}
   "black"   {:500 "000000"}
   "transparent" {:500 nil}})

;; Text sizes mapping to Excel font sizes
(def ^:private text-sizes
  {"text-xs"   8
   "text-sm"   10
   "text-base" 12
   "text-lg"   14
   "text-xl"   18
   "text-2xl"  24
   "text-3xl"  30
   "text-4xl"  36})

;; Border styles
(def ^:private border-styles
  {"border-solid"  "thin"
   "border-dashed" "dashed"
   "border-dotted" "dotted"
   "border-double" "double"
   "border-2"      "medium"
   "border-4"      "thick"})

(defn- parse-color
  "Parse a Tailwind color class like 'blue-500' or 'red' into hex.
   Returns nil for invalid colors."
  [color-str]
  (if-let [[_ color shade] (re-matches #"(\w+)-?(\d+)?" color-str)]
    (when-let [color-map (get colors color)]
      (get color-map (keyword (or shade "500"))))
    nil))

(defn- parse-class
  "Parse a single Tailwind class into a style map fragment."
  [class-str]
  (cond
    ;; Background color: bg-{color}-{shade}
    (str/starts-with? class-str "bg-")
    (when-let [hex (parse-color (subs class-str 3))]
      {:fill hex})

    ;; Text color: text-{color}-{shade}
    (and (str/starts-with? class-str "text-")
         (not (contains? text-sizes class-str))
         (not (#{"text-left" "text-center" "text-right"} class-str)))
    (when-let [hex (parse-color (subs class-str 5))]
      {:font-color hex})

    ;; Text size
    (contains? text-sizes class-str)
    {:font-size (get text-sizes class-str)}

    ;; Font weight
    (= class-str "font-bold")
    {:bold true}

    (= class-str "font-normal")
    {:bold false}

    ;; Font style
    (= class-str "italic")
    {:italic true}

    (= class-str "underline")
    {:underline true}

    ;; Text alignment
    (= class-str "text-left")
    {:align :left}

    (= class-str "text-center")
    {:align :center}

    (= class-str "text-right")
    {:align :right}

    ;; Vertical alignment
    (= class-str "align-top")
    {:valign :top}

    (= class-str "align-middle")
    {:valign :center}

    (= class-str "align-bottom")
    {:valign :bottom}

    ;; Border all sides
    (= class-str "border")
    {:border-left "thin" :border-right "thin"
     :border-top "thin" :border-bottom "thin"}

    ;; Border specific sides
    (= class-str "border-t")
    {:border-top "thin"}

    (= class-str "border-b")
    {:border-bottom "thin"}

    (= class-str "border-l")
    {:border-left "thin"}

    (= class-str "border-r")
    {:border-right "thin"}

    ;; Border style
    (contains? border-styles class-str)
    (let [style (get border-styles class-str)]
      {:border-left style :border-right style
       :border-top style :border-bottom style})

    ;; Border color: border-{color}-{shade}
    (str/starts-with? class-str "border-")
    (let [rest (subs class-str 7)]
      (when-let [hex (parse-color rest)]
        {:border-color hex}))

    ;; Text wrap
    (= class-str "whitespace-normal")
    {:wrap true}

    (= class-str "whitespace-nowrap")
    {:wrap false}

    :else nil))

(defn- parse-classes
  "Parse multiple Tailwind classes into a combined style map."
  [classes]
  (->> (str/split classes #"\.")
       (filter seq)
       (map parse-class)
       (filter some?)
       (apply merge)))

(defn- parse-cell-ref
  "Parse a cell reference like 'A1', 'AA23', etc.
   Returns {:col \"A\" :row 1} or nil if invalid."
  [ref-str]
  (when-let [[_ col row] (re-matches #"([A-Z]+)(\d+)" (str/upper-case ref-str))]
    {:col col :row (parse-long row)}))

(defn- parse-selector
  "Parse a selector keyword like :A1.bg-blue-500 or :Sheet1/A1.text-white
   Returns {:sheet \"Sheet1\" :target {:type :cell :col \"A\" :row 1} :styles {...}}"
  [selector]
  (let [;; Use keyword namespace for sheet name (e.g., :Sheet1/A1 has namespace "Sheet1")
        sheet-part (namespace selector)
        s (name selector)
        ;; Split target from classes
        [target & classes] (str/split s #"\.")
        target-upper (str/upper-case target)
        styles (parse-classes (str/join "." classes))]
    {:sheet sheet-part
     :styles styles
     :target
     (cond
       ;; Cell reference like A1, AA23
       (re-matches #"[A-Z]+\d+" target-upper)
       (merge {:type :cell} (parse-cell-ref target-upper))

       ;; Row number like 5, 23
       (re-matches #"\d+" target)
       {:type :row :row (parse-long target)}

       ;; Column letter(s) like A, AA, BC
       (re-matches #"[A-Z]+" target-upper)
       {:type :column :col target-upper}

       ;; Range like A1:C3
       (re-matches #"[A-Z]+\d+:[A-Z]+\d+" target-upper)
       (let [[start end] (str/split target-upper #":")]
         {:type :range
          :start (parse-cell-ref start)
          :end (parse-cell-ref end)})

       :else
       {:type :unknown :raw target})}))

;; ============= Excel XML Generation =============

(def ^:private xml-header "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")

(def ^:private xlsx-ns
  {:xmlns "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
   :xmlns:r "http://schemas.openxmlformats.org/officeDocument/2006/relationships"})

(defn- style->font-xml
  "Convert style map to Excel font XML element."
  [{:keys [font-size bold italic underline font-color]}]
  [:font
   (when font-size [:sz {:val font-size}])
   (when bold [:b])
   (when italic [:i])
   (when underline [:u])
   (if font-color
     [:color {:rgb (str "FF" font-color)}]
     [:color {:theme "1"}])
   [:name {:val "Calibri"}]
   [:family {:val "2"}]])

(defn- style->fill-xml
  "Convert style map to Excel fill XML element."
  [{:keys [fill]}]
  (if fill
    [:fill
     [:patternFill {:patternType "solid"}
      [:fgColor {:rgb (str "FF" fill)}]
      [:bgColor {:indexed "64"}]]]
    [:fill
     [:patternFill {:patternType "none"}]]))

(defn- style->border-xml
  "Convert style map to Excel border XML element."
  [{:keys [border-left border-right border-top border-bottom border-color]}]
  (let [make-border (fn [style]
                      (if style
                        {:style style}
                        {}))
        color-elem (when border-color
                     [:color {:rgb (str "FF" border-color)}])]
    [:border
     [:left (make-border border-left) color-elem]
     [:right (make-border border-right) color-elem]
     [:top (make-border border-top) color-elem]
     [:bottom (make-border border-bottom) color-elem]
     [:diagonal]]))

(defn- style->xf-xml
  "Convert style map to Excel cellXf XML element."
  [style font-id fill-id border-id]
  (let [{:keys [align valign wrap]} style
        has-alignment (or align valign wrap)]
    [:xf {:numFmtId "0"
          :fontId (str font-id)
          :fillId (str fill-id)
          :borderId (str border-id)
          :xfId "0"
          :applyFont (if (some style [:bold :italic :underline :font-size :font-color]) "1" "0")
          :applyFill (if (:fill style) "1" "0")
          :applyBorder (if (some style [:border-left :border-right :border-top :border-bottom]) "1" "0")
          :applyAlignment (if has-alignment "1" "0")}
     (when has-alignment
       [:alignment
        (cond-> {}
          align (assoc :horizontal (name align))
          valign (assoc :vertical (name valign))
          wrap (assoc :wrapText "1"))])]))

(defn- collect-unique-styles
  "Collect all unique styles from cells and assign indices."
  [cells]
  (let [all-styles (->> cells
                        (map :styles)
                        (filter seq)
                        distinct
                        vec)]
    ;; Return a map from style to index (0 is reserved for default)
    (into {} (map-indexed (fn [i s] [s (inc i)]) all-styles))))

(defn- generate-styles-xml
  "Generate the complete styles.xml content."
  [style-index]
  (let [styles (keys style-index)
        ;; Default + custom fonts
        fonts (cons (style->font-xml {}) (map style->font-xml styles))
        ;; Default fills (none, gray125) + custom
        fills (concat [[:fill [:patternFill {:patternType "none"}]]
                       [:fill [:patternFill {:patternType "gray125"}]]]
                      (map style->fill-xml styles))
        ;; Default border + custom
        borders (cons (style->border-xml {}) (map style->border-xml styles))
        ;; Cell formats
        xfs (cons [:xf {:numFmtId "0" :fontId "0" :fillId "0" :borderId "0" :xfId "0"}]
                  (map-indexed (fn [i s]
                                 (style->xf-xml s (inc i) (+ 2 i) (inc i)))
                               styles))]
    (str xml-header
         (hc/html
          [:styleSheet {:xmlns "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
           (into [:fonts {:count (str (count fonts))}] fonts)
           (into [:fills {:count (str (count fills))}] fills)
           (into [:borders {:count (str (count borders))}] borders)
           [:cellStyleXfs {:count "1"}
            [:xf {:numFmtId "0" :fontId "0" :fillId "0" :borderId "0"}]]
           (into [:cellXfs {:count (str (count xfs))}] xfs)
           [:cellStyles {:count "1"}
            [:cellStyle {:name "Normal" :xfId "0" :builtinId "0"}]]]))))

(defn- cell-type-and-value
  "Determine cell type and formatted value for Excel."
  [value]
  (cond
    (string? value) ["inlineStr" [:is [:t value]]]
    (number? value) ["n" [:v value]]
    (boolean? value) ["b" [:v (if value "1" "0")]]
    (instance? java.time.LocalDate value)
    ["n" [:v (.between java.time.temporal.ChronoUnit/DAYS
                       (java.time.LocalDate/of 1899 java.time.Month/DECEMBER 30)
                       value)]]
    :else ["inlineStr" [:is [:t (str value)]]]))

(defn- generate-cell-xml
  "Generate XML for a single cell."
  [col row value style-id]
  (let [[t v] (cell-type-and-value value)]
    [:c (cond-> {:r (str col row) :t t}
          (and style-id (pos? style-id)) (assoc :s (str style-id)))
     v]))

(defn- generate-sheet-xml
  "Generate worksheet XML from cell data."
  [cells style-index]
  (let [;; Group cells by row
        by-row (group-by #(get-in % [:target :row]) cells)
        row-nums (sort (keys by-row))]
    (str xml-header
         (hc/html
          [:worksheet xlsx-ns
           [:sheetData
            (for [row-num row-nums
                  :let [row-cells (get by-row row-num)]]
              [:row {:r (str row-num)}
               (for [{:keys [target styles value]} (sort-by #(get-in % [:target :col]) row-cells)
                     :let [style-id (get style-index styles 0)]]
                 (generate-cell-xml (:col target) row-num value style-id))])]]))))

(defn- generate-workbook-xml
  "Generate workbook.xml content."
  [sheet-names]
  (str xml-header
       (hc/html
        [:workbook xlsx-ns
         (into [:sheets]
               (map-indexed (fn [i name]
                              [:sheet {:name name
                                       :sheetId (str (inc i))
                                       :r:id (str "rId" (inc i))}])
                            sheet-names))])))

(defn- generate-workbook-rels
  "Generate xl/_rels/workbook.xml.rels content."
  [sheet-count]
  (str xml-header
       (hc/html
        (into [:Relationships {:xmlns "http://schemas.openxmlformats.org/package/2006/relationships"}]
              (concat
               (for [i (range sheet-count)]
                 [:Relationship {:Id (str "rId" (inc i))
                                 :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                                 :Target (str "worksheets/sheet" (inc i) ".xml")}])
               [[:Relationship {:Id (str "rId" (inc sheet-count))
                                :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
                                :Target "styles.xml"}]])))))

(defn- generate-content-types
  "Generate [Content_Types].xml content."
  [sheet-count]
  (str xml-header
       (hc/html
        (into [:Types {:xmlns "http://schemas.openxmlformats.org/package/2006/content-types"}
               [:Default {:Extension "rels"
                          :ContentType "application/vnd.openxmlformats-package.relationships+xml"}]
               [:Default {:Extension "xml" :ContentType "application/xml"}]
               [:Override {:PartName "/xl/workbook.xml"
                           :ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"}]
               [:Override {:PartName "/xl/styles.xml"
                           :ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"}]]
              (for [i (range sheet-count)]
                [:Override {:PartName (str "/xl/worksheets/sheet" (inc i) ".xml")
                            :ContentType "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"}])))))

(defn- generate-root-rels
  "Generate _rels/.rels content."
  []
  (str xml-header
       (hc/html
        [:Relationships {:xmlns "http://schemas.openxmlformats.org/package/2006/relationships"}
         [:Relationship {:Id "rId1"
                         :Type "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                         :Target "xl/workbook.xml"}]])))

(defn- write-zip-entry
  "Write a string entry to a ZipOutputStream."
  [^ZipOutputStream zos ^String path ^String content]
  (.putNextEntry zos (ZipEntry. path))
  (.write zos (.getBytes content "UTF-8"))
  (.closeEntry zos))

(defn- expand-cell
  "Expand a styled element into individual cell entries.
   Handles single cells, rows, columns, and ranges."
  [[selector value]]
  (let [{:keys [sheet target styles]} (parse-selector selector)]
    (case (:type target)
      :cell
      [{:sheet sheet :target target :styles styles :value value}]

      :row
      (if (map? value)
        ;; {:A "val1" :B "val2"}
        (for [[col v] value]
          {:sheet sheet
           :target {:type :cell :col (name col) :row (:row target)}
           :styles styles
           :value v})
        ;; Vector of values
        (map-indexed (fn [i v]
                       {:sheet sheet
                        :target {:type :cell
                                 :col (core/number->column-letter (inc i))
                                 :row (:row target)}
                        :styles styles
                        :value v})
                     value))

      :column
      (if (map? value)
        ;; {:1 "val1" :2 "val2"}
        (for [[row v] value]
          {:sheet sheet
           :target {:type :cell :col (:col target) :row (parse-long (name row))}
           :styles styles
           :value v})
        ;; Vector of values
        (map-indexed (fn [i v]
                       {:sheet sheet
                        :target {:type :cell :col (:col target) :row (inc i)}
                        :styles styles
                        :value v})
                     value))

      :range
      (let [{:keys [start end]} target
            start-col-num (core/column-letter->number (:col start))
            end-col-num (core/column-letter->number (:col end))]
        (for [row (range (:row start) (inc (:row end)))
              col-num (range start-col-num (inc end-col-num))
              :let [col (core/number->column-letter col-num)
                    row-idx (- row (:row start))
                    col-idx (- col-num start-col-num)
                    v (get-in value [row-idx col-idx])]]
          {:sheet sheet
           :target {:type :cell :col col :row row}
           :styles styles
           :value v}))

      ;; Unknown/unsupported
      [])))

(defn create-styled-xlsx
  "Create a styled Excel spreadsheet using Tailwind-like CSS classes.

   Usage:
     (create-styled-xlsx \"output.xlsx\"
       [[:A1.bg-blue-500.text-white.font-bold \"Header\"]
        [:A2.border \"Data\"]
        [:B1:B3.bg-gray-100 [[\"Col B\"] [1] [2]]]])

   Supported classes:
     bg-{color}-{shade}  - Background color
     text-{color}-{shade} - Text color
     text-{size}         - Font size (xs, sm, base, lg, xl, 2xl)
     font-bold           - Bold text
     italic              - Italic text
     underline           - Underlined text
     text-left/center/right - Horizontal alignment
     border              - All borders
     border-t/b/l/r      - Individual borders
     border-{color}-{shade} - Border color
     border-dashed/dotted/double - Border style"
  [file-path data]
  (let [;; Expand all cells
        all-cells (mapcat expand-cell data)
        ;; Group by sheet (nil = default "Sheet1")
        by-sheet (group-by :sheet all-cells)
        sheet-names (or (seq (filter some? (keys by-sheet)))
                        ["Sheet1"])
        ;; Ensure we have a default sheet
        by-sheet (if (contains? by-sheet nil)
                   (assoc (dissoc by-sheet nil)
                          (first sheet-names)
                          (get by-sheet nil))
                   by-sheet)
        ;; Collect all unique styles across all sheets
        style-index (collect-unique-styles all-cells)]

    (io/make-parents file-path)
    (with-open [fos (FileOutputStream. ^String file-path)
                zos (ZipOutputStream. fos)]
      ;; Write worksheets
      (doseq [[i sheet-name] (map-indexed vector sheet-names)
              :let [sheet-cells (get by-sheet sheet-name [])]]
        (write-zip-entry zos
                         (str "xl/worksheets/sheet" (inc i) ".xml")
                         (generate-sheet-xml sheet-cells style-index)))

      ;; Write styles
      (write-zip-entry zos "xl/styles.xml" (generate-styles-xml style-index))

      ;; Write workbook
      (write-zip-entry zos "xl/workbook.xml" (generate-workbook-xml sheet-names))

      ;; Write relationships
      (write-zip-entry zos "xl/_rels/workbook.xml.rels"
                       (generate-workbook-rels (count sheet-names)))
      (write-zip-entry zos "_rels/.rels" (generate-root-rels))

      ;; Write content types
      (write-zip-entry zos "[Content_Types].xml"
                       (generate-content-types (count sheet-names))))))

;; ============= Convenience macros/functions =============

(defn styled-row
  "Create a styled row with uniform styling.
   (styled-row 1 :bg-blue-500.font-bold [\"A\" \"B\" \"C\"])"
  [row-num style-kw values]
  [(keyword (str row-num "." (name style-kw))) values])

(defn styled-col
  "Create a styled column with uniform styling.
   (styled-col :A :border.text-center {1 \"Header\" 2 100 3 200})"
  [col style-kw values]
  [(keyword (str (name col) "." (name style-kw))) values])
