(ns core-test
  (:require [clojure.java.io :as io]
            [clojure.set :refer [intersection]]
            [clojure.test :refer [deftest is run-tests testing]]
            [bb-excel.core :refer [get-sheets get-sheet-names get-sheet
                                   get-range create-xlsx
                                   get-table-names get-table]]
            [bb-excel.styled :as styled]
            [malli.core :as malli]
            [malli.generator :as mg])
  (:import (clojure.lang ExceptionInfo)
           [java.util.zip ZipFile]))

(declare thrown-with-msg?) ;; Workaround from https://github.com/cursive-ide/cursive/issues/238

(deftest zipfile-or-nil-test
  (let [zipfile-or-nil #'bb-excel.core/get-zipfile]
    (let [file (io/file "test/data/simple.xlsx")]
      (is (instance? ZipFile (zipfile-or-nil file))))
    (let [filepath "test/data/simple.xlsx"]
      (is (instance? ZipFile (zipfile-or-nil filepath))))
    (is (thrown-with-msg? ExceptionInfo #"Could not open 'invalid-file-path'! File does not exist."
                          (zipfile-or-nil "invalid-file-path")))
    (is (thrown-with-msg? ExceptionInfo #"Could not open ':invalid-type'! Argument should be string or file."
                          (zipfile-or-nil :invalid-type)))))

(deftest get-sheet-names-test
  (testing "Get Sheet Names"
    (is (= '({:name "Sheet1" :idx 1}
             {:name "Shows" :idx 2})
           (get-sheet-names "test/data/simple.xlsx")))
    (is (thrown-with-msg? ExceptionInfo #"Could not open 'missingfile.xlsx'! File does not exist."
                          (get-sheet-names "missingfile.xlsx")))
    (is (thrown-with-msg? ExceptionInfo #"Could not open 'null'! Argument should be string or file."
                          (get-sheet-names nil)))))

(deftest get-sheets-test
  (testing "Get Sheets"
    (is (= '({:name "Sheet1" :idx 1,
              :sheet ({:_r 1, :A "FirstName", :B "LastName", :C "DateOfBirth", :D "Show", :E "Votes"}
                      {:_r 2, :A "Jack", :B "Bean", :C "04/20/1979", :D "Breaking Bad", :E "1325"}
                      {:_r 3, :A "Mary", :B "Smith", :C "05/15/1991", :D "House M.D", :E "435"}
                      {:_r 4, :A "Todd", :B "Green", :C "12/31/1999", :D "La Femme Nikita", :E "80"})}
             {:name "Shows" :idx 2,
              :sheet ({:_r 1, :A "Rank", :B "TV Show"}
                      {:_r 2, :A "1", :B "Sesame Street"}
                      {:_r 3, :A "2", :B "La Femme Nikita"}
                      {:_r 4, :A "3", :B "House M.D"}
                      {:_r 5, :A "4", :B "Breaking Bad"})})
           (get-sheets "test/data/simple.xlsx")))
    (is (thrown-with-msg? ExceptionInfo #"Could not open 'missingfile.xlsx'! File does not exist."
                          (get-sheet-names "missingfile.xlsx")))
    (is (thrown-with-msg? ExceptionInfo #"Could not open 'null'! Argument should be string or file."
                          (get-sheets nil)))))

(deftest get-range-test
  (testing "Get Sheet Range"
    (is (= '({:_r 1, :A "FirstName", :B "LastName"}
             {:_r 2, :A "Jack", :B "Bean"})
           (get-range (get-sheet "test/data/simple.xlsx" "Sheet1") "A1:B2")))
    (is (thrown-with-msg? ExceptionInfo #"Could not open 'missingfile.xlsx'! File does not exist."
                          (get-sheet-names "missingfile.xlsx")))
    (is (thrown-with-msg? ExceptionInfo #"Could not open 'null'! Argument should be string or file."
                          (get-sheet-names nil)))
    (is (= '({:_r 10 :A "9" :B "TextData"})
           (get-range (get-sheet "test/data/Types.xlsx" "Sheet1") "A10:B10")))))

(deftest corner-cases-test
  (testing "Without shared files"
    (is (= '({:_r 1, :A 1})
           (get-sheet "test/data/without_sharedfiles.xlsx" 1)))))

;; Issue #17: get-sheet by name loads wrong data when Excel file has deleted sheets
;; https://github.com/kbosompem/bb-excel/issues/17
(deftest deleted-sheets-test
  (testing "Sheet names with non-sequential IDs (simulating deleted sheets)"
    ;; The file has sheetIds 1, 4, 5 but the relationship IDs map correctly
    (is (= [{:name "Users" :idx 1}
            {:name "Communities" :idx 4}
            {:name "Zones" :idx 5}]
           (get-sheet-names "test/data/deleted_sheets.xlsx"))))

  (testing "Loading sheets by name correctly maps to actual worksheet files"
    ;; Users should load from sheet1.xml (via rId1)
    (let [users (get-sheet "test/data/deleted_sheets.xlsx" "Users")]
      (is (= #{:A :B :C :_r} (set (keys (first users)))))
      (is (= "user_id" (:A (first users)))))

    ;; Communities should load from sheet2.xml (via rId2), not sheet4.xml
    (let [communities (get-sheet "test/data/deleted_sheets.xlsx" "Communities")]
      (is (= "community_id" (:A (first communities)))))

    ;; Zones should load from sheet3.xml (via rId3), not sheet5.xml
    (let [zones (get-sheet "test/data/deleted_sheets.xlsx" "Zones")]
      (is (= "zone_id" (:A (first zones))))))

  (testing "Loading sheets by positional index (1-based)"
    ;; Position 1 = Users (first sheet in list)
    (is (= "user_id" (:A (first (get-sheet "test/data/deleted_sheets.xlsx" 1)))))
    ;; Position 2 = Communities (second sheet in list)
    (is (= "community_id" (:A (first (get-sheet "test/data/deleted_sheets.xlsx" 2)))))
    ;; Position 3 = Zones (third sheet in list)
    (is (= "zone_id" (:A (first (get-sheet "test/data/deleted_sheets.xlsx" 3)))))))

;; Issue #18: Header columns randomly missing when parsing xlsx with get-sheet
;; https://github.com/kbosompem/bb-excel/issues/18
(deftest missing-cell-refs-test
  (testing "Cells without r attribute are assigned sequential column letters"
    ;; File where all cells lack the r attribute
    (let [data (get-sheet "test/data/no_cell_refs.xlsx" "NoRefs")]
      ;; Row 1 should have columns A through E
      (is (= #{:A :B :C :D :E :_r} (set (keys (first data)))))
      (is (= "col_a" (:A (first data))))
      (is (= "col_b" (:B (first data))))
      (is (= "col_c" (:C (first data))))
      (is (= "col_d" (:D (first data))))
      (is (= "col_e" (:E (first data))))

      ;; Row 2 should also have columns A through E
      (is (= "val_a1" (:A (second data))))
      (is (= "val_e1" (:E (second data))))))

  (testing "Mixed cells with and without r attributes"
    (let [data (get-sheet "test/data/mixed_refs.xlsx" "MixedRefs")]
      ;; Row 1: A1, B1 have refs, then C, D, E continue sequentially
      (is (= "header_a" (:A (first data))))
      (is (= "header_b" (:B (first data))))
      (is (= "header_c" (:C (first data))))
      (is (= "header_d" (:D (first data))))
      (is (= "header_e" (:E (first data))))

      ;; Row 2: A2 has ref, gap (no B), C2 has ref, D, E continue from C
      ;; Expected: A=a_val, B=nil, C=c_val, D=d_val, E=e_val
      (is (= "a_val" (:A (second data))))
      (is (nil? (:B (second data)))) ;; Gap - B is missing
      (is (= "c_val" (:C (second data))))
      (is (= "d_val" (:D (second data))))
      (is (= "e_val" (:E (second data))))

      ;; Row 3: All cells without refs, should be A through E
      (is (= "row3_a" (:A (nth data 2))))
      (is (= "row3_e" (:E (nth data 2))))))

  (testing "Header mode works correctly with missing cell refs"
    (let [data (get-sheet "test/data/no_cell_refs.xlsx" "NoRefs" {:hdr true :row 1})]
      ;; Headers should be col_a, col_b, etc.
      (is (= #{"col_a" "col_b" "col_c" "col_d" "col_e" :_r} (set (keys (first data)))))
      (is (= "val_a1" (get (first data) "col_a")))
      (is (= "val_e1" (get (first data) "col_e"))))))

(deftest create-xlsx-test
  (testing "Creating an Excel Spreadsheet"
    (is (= #{{:A "2", :B "Two", :C "Mienu"} {:A "1", :B "One", :C "Baako"} {:A "3", :B "Three", :C "Miensa"}}
           (let [d [{:name "TestSheet"
                     :sheet [{:A "1" :B "One" :C "Baako"}
                             {:A "2" :B "Two" :C "Mienu"}
                             {:A "3" :B "Three" :C "Miensa"}]}]
                 _ (create-xlsx "zomb.xlsx" d)
                 xs (get-sheets "zomb.xlsx")
                 data  (-> xs
                           first
                           (dissoc :idx)
                           :sheet
                           (->> (map #(dissoc % :_r))))
                 ins (clojure.set/intersection (set (:sheet (first d))) (set data))]
             ins)))))

;; Experimental: Tailwind-styled xlsx creation
(deftest styled-xlsx-test
  (testing "Parse Tailwind classes"
    (is (= {:fill "3B82F6" :font-color "FFFFFF" :bold true}
           (#'styled/parse-classes "bg-blue-500.text-white.font-bold")))
    (is (= {:align :right}
           (#'styled/parse-classes "text-right")))
    (is (= {:border-left "thin" :border-right "thin"
            :border-top "thin" :border-bottom "thin"}
           (#'styled/parse-classes "border"))))

  (testing "Parse cell selectors"
    (let [parsed (#'styled/parse-selector :A1.bg-red-500)]
      (is (= :cell (get-in parsed [:target :type])))
      (is (= "A" (get-in parsed [:target :col])))
      (is (= 1 (get-in parsed [:target :row])))
      (is (= "EF4444" (get-in parsed [:styles :fill]))))

    (let [parsed (#'styled/parse-selector :5.font-bold)]
      (is (= :row (get-in parsed [:target :type])))
      (is (= 5 (get-in parsed [:target :row]))))

    (let [parsed (#'styled/parse-selector :AA.border)]
      (is (= :column (get-in parsed [:target :type])))
      (is (= "AA" (get-in parsed [:target :col]))))

    (let [parsed (#'styled/parse-selector :Sheet1/B2.italic)]
      (is (= "Sheet1" (:sheet parsed)))
      (is (= "B" (get-in parsed [:target :col])))))

  (testing "Create styled xlsx and read back data"
    (styled/create-styled-xlsx "styled_test.xlsx"
                               [[:A1.bg-blue-500.font-bold "Header"]
                                [:A2 "Data"]
                                [:B1.text-right 100]
                                [:B2 200]])
    (let [data (get-sheet "styled_test.xlsx" 1)]
      (is (= 2 (count data)))
      (is (= "Header" (:A (first data))))
      (is (= 100 (:B (first data))))
      (is (= "Data" (:A (second data))))
      (is (= 200 (:B (second data))))))

  (testing "Row styling applies to all cells"
    (styled/create-styled-xlsx "styled_row_test.xlsx"
                               [[:1.bg-gray-100 ["A" "B" "C"]]
                                [:2 ["D" "E" "F"]]])
    (let [data (get-sheet "styled_row_test.xlsx" 1)]
      (is (= "A" (:A (first data))))
      (is (= "C" (:C (first data))))
      (is (= "D" (:A (second data))))))

  (testing "Range styling"
    (styled/create-styled-xlsx "styled_range_test.xlsx"
                               [[:A1:B2.border [[1 2] [3 4]]]])
    (let [data (get-sheet "styled_range_test.xlsx" 1)]
      (is (= 1 (:A (first data))))
      (is (= 2 (:B (first data))))
      (is (= 3 (:A (second data))))
      (is (= 4 (:B (second data)))))))

;; Issue #14: Add support for reading Excel tables
(deftest excel-tables-test
  (testing "List all tables in a workbook"
    (let [tables (get-table-names "test/data/tables.xlsx")]
      (is (= 2 (count tables)))
      (is (= #{"PeopleTable" "ProductsTable"} (set (map :name tables))))))

  (testing "PeopleTable metadata"
    (let [tables (get-table-names "test/data/tables.xlsx")
          people (first (filter #(= "PeopleTable" (:name %)) tables))]
      (is (= "People" (:sheet people)))
      (is (= ["Name" "Age" "City" "Score"] (:columns people)))
      (is (= "A1:D5" (:ref people)))))

  (testing "ProductsTable metadata"
    (let [tables (get-table-names "test/data/tables.xlsx")
          products (first (filter #(= "ProductsTable" (:name %)) tables))]
      (is (= "Products" (:sheet products)))
      (is (= ["ProductName" "Price" "Category"] (:columns products)))
      (is (= "A1:C4" (:ref products)))))

  (testing "Get PeopleTable data"
    (let [{:keys [name data]} (get-table "test/data/tables.xlsx" "PeopleTable")]
      (is (= "PeopleTable" name))
      (is (= 4 (count data)))
      (is (= "Alice" (get (first data) "Name")))
      (is (= 30 (get (first data) "Age")))
      (is (= "New York" (get (first data) "City")))
      (is (= "Bob" (get (second data) "Name")))
      (is (= 25 (get (second data) "Age")))))

  (testing "Get ProductsTable data"
    (let [{:keys [name data]} (get-table "test/data/tables.xlsx" "ProductsTable")]
      (is (= "ProductsTable" name))
      (is (= 3 (count data)))
      (is (= "Widget" (get (first data) "ProductName")))
      (is (= "Electronics" (get (first data) "Category")))
      (is (= "Thingamajig" (get (nth data 2) "ProductName")))
      (is (= "Tools" (get (nth data 2) "Category")))))

  (testing "Missing table throws exception"
    (is (thrown-with-msg? ExceptionInfo #"Could not find table 'NonExistentTable'!"
                          (get-table "test/data/tables.xlsx" "NonExistentTable"))))

  (testing "Workbook without tables returns empty list"
    (is (= [] (get-table-names "test/data/simple.xlsx")))))

;; Issue #15: Add support for writing excel tables
(deftest write-table-xlsx-test
  (testing "Create table with :table true and verify via get-table"
    (create-xlsx "tables_test.xlsx"
                 [{:name "People"
                   :table true
                   :sheet [{"Name" "Alice" "Age" 30 "City" "New York"}
                           {"Name" "Bob"   "Age" 25 "City" "London"}
                           {"Name" "Carol" "Age" 35 "City" "Toronto"}]}])
    (let [tables (get-table-names "tables_test.xlsx")]
      (is (= 1 (count tables)))
      (is (= "Table1" (:name (first tables))))
      (is (= "People" (:sheet (first tables))))
      (is (= ["Name" "Age" "City"] (:columns (first tables))))
      (is (= "A1:C4" (:ref (first tables)))))
    (let [{:keys [name data]} (get-table "tables_test.xlsx" "Table1")]
      (is (= "Table1" name))
      (is (= 3 (count data)))
      (is (= "Alice" (get (first data) "Name")))
      (is (= 30 (get (first data) "Age")))
      (is (= "London" (get (second data) "City")))))

  (testing "Create table with custom name and style"
    (create-xlsx "tables_test.xlsx"
                 [{:name "Products"
                   :table {:name "SalesData" :style "TableStyleLight1"}
                   :sheet [{"Product" "Widget" "Price" 9.99}
                           {"Product" "Gadget" "Price" 19.99}]}])
    (let [tables (get-table-names "tables_test.xlsx")]
      (is (= 1 (count tables)))
      (is (= "SalesData" (:name (first tables))))
      (is (= ["Product" "Price"] (:columns (first tables)))))
    (let [{:keys [data]} (get-table "tables_test.xlsx" "SalesData")]
      (is (= 2 (count data)))
      (is (= "Widget" (get (first data) "Product")))))

  (testing "Mixed sheets — only table sheet has table metadata"
    (create-xlsx "tables_test.xlsx"
                 [{:name "Plain"
                   :sheet [{"X" "a" "Y" "b"}
                           {"X" "c" "Y" "d"}]}
                  {:name "WithTable"
                   :table {:name "MyTable"}
                   :sheet [{"Col1" "foo" "Col2" "bar"}
                           {"Col1" "baz" "Col2" "qux"}]}])
    (let [tables (get-table-names "tables_test.xlsx")]
      (is (= 1 (count tables)))
      (is (= "MyTable" (:name (first tables))))
      (is (= "WithTable" (:sheet (first tables)))))
    ;; Verify the ZIP has table files only for the table sheet
    (let [zf (ZipFile. "tables_test.xlsx")]
      (is (some? (.getEntry zf "xl/tables/table1.xml")))
      (is (some? (.getEntry zf "xl/worksheets/_rels/sheet2.xml.rels")))
      (is (nil?  (.getEntry zf "xl/worksheets/_rels/sheet1.xml.rels")))
      (.close zf)))

  (testing "Multiple table sheets get distinct table IDs"
    (create-xlsx "tables_test.xlsx"
                 [{:name "Sheet1"
                   :table {:name "TableA"}
                   :sheet [{"ID" 1 "Val" "x"}]}
                  {:name "Sheet2"
                   :table {:name "TableB"}
                   :sheet [{"ID" 2 "Val" "y"}]}])
    (let [tables (get-table-names "tables_test.xlsx")
          names  (set (map :name tables))]
      (is (= 2 (count tables)))
      (is (= #{"TableA" "TableB"} names)))
    (let [zf (ZipFile. "tables_test.xlsx")]
      (is (some? (.getEntry zf "xl/tables/table1.xml")))
      (is (some? (.getEntry zf "xl/tables/table2.xml")))
      (.close zf)))

  (testing "Existing create-xlsx still works without :table"
    (create-xlsx "tables_test.xlsx"
                 [{:name "TestSheet"
                   :sheet [{"A" "1" "B" "One"}
                           {"A" "2" "B" "Two"}]}])
    (is (= [] (get-table-names "tables_test.xlsx")))
    (let [zf (ZipFile. "tables_test.xlsx")]
      (is (nil? (.getEntry zf "xl/tables/table1.xml")))
      (.close zf))))

(comment
  (run-tests)

  (create-xlsx "sample.xlsx"    [{:name "TestSheet"
                                  :sheet [{:A "1" :B "One" :C "Baako"}
                                          {:A "2" :B "Two" :C "Mienu"}
                                          {:A "3" :B "Three" :C "Miensa"}]}])
   ;  To validate the data was saved
  (clojure.pprint/print-table
   (get-sheet "sample.xlsx" "TestSheet" {:hdr true}))

  (get-sheet "test/data/simple.xlsx" "Shows" {:hdr true :row 1})

  (create-xlsx "output/kay.xlsx" [{:name "TVShows"
                                   :sheet [{"Rank" "1", "TV Show" "Sesame Street"}
                                           {"Rank" "2", "TV Show" "La Femme Nikita"}
                                           {"Rank" "3", "TV Show" "House M.D"}
                                           {"Rank" "4", "TV Show" "Breaking Bad"}]}
                                  {:name "Shows-1"
                                   :sheet [{"Rank" "1", "TV Show" "1Sesame Street"}
                                           {"Rank" "2", "TV Show" "1La Femme Nikita"}
                                           {"Rank" "3", "TV Show" "1House M.D"}
                                           {"Rank" "4", "TV Show" "1Breaking Bad"}]}
                                  {:name "Shows-2"
                                   :sheet [{"Rank" "1", "TV Show" "2Sesame Street"}
                                           {"Rank" "2", "TV Show" "2La Femme Nikita"}
                                           {"Rank" "3", "TV Show" "2House M.D"}
                                           {"Rank" "4", "TV Show" (java.time.LocalDate/now)}]}])

  (create-xlsx "output/ghana.xlsx" [{:name "mama"
                                     :sheet [["Col A" "Col B" "Col C" "Col D" "Col E"]
                                             [\1 2 3 4 5]
                                             [1 \2 3 4 (java.time.LocalDate/now)]
                                             [\a \b \c \d \e]]}])

  (create-xlsx "output/ghana.xlsx" [[1 2 3 4 5]
                                    [1 2 3 4 5]
                                    [\a \b \c \d \e]])

  (get-sheet "output/kay.xlsx" "TVShows" {:hdr true :row 1})

  (get-sheet "output/sample.xlsx" "TestSheet" {:hdr true :row 1 :fxn (comp keyword str)})

  (def MSheet [:vector {:min 1 :max 4} map?])
  (def VSheet [:vector {:min 1 :max 4} vector?])
  (def Workbook [:vector [:map
                          [:name :string]
                          [:cmap {:optional true} map?]
                          [:idx  {:optional true} :int]
                          [:sheet  [:or MSheet VSheet]]]])

  (create-xlsx "sosket.xlsx" (malli.generator/generate Workbook))
  (create-xlsx "maga.xlsx" [{:name "2R6a325retiLS5IvCtV", :sheet [[]]}])
  #{})
