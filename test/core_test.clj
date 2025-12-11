(ns core-test
  (:require [clojure.java.io :as io]
            [clojure.set :refer [intersection]]
            [clojure.test :refer [deftest is run-tests testing]]
            [bb-excel.core :refer [get-sheets get-sheet-names get-sheet
                                   get-range create-xlsx]]
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
