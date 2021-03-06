(ns core-test
  (:require [clojure.test :refer [deftest is testing run-tests]]
            [bb-excel.core :refer [get-sheets get-sheet-names get-sheet
                                   get-range get-row get-col get-cells crange]]))

(deftest get-sheet-names-test
  (testing "Get Sheet Names"
    (is (= '({:id "rId1", :name "Sheet1", :sheetId "1", :idx 1})
           (get-sheet-names "test/data/simple.xlsx")))
    (is (nil? (get-sheet-names "missingfile.xlsx"))
        "File does not exist. Should return null.")
    (is (nil? (get-sheet-names nil))
        "Filename was not passed in")))

(deftest get-sheets-test
  (testing "Get Sheets"
    (is (= '({:id "rId1",
              :name "Sheet1",
              :sheetId "1",
              :idx 1,
              :sheet ({:_r 1, :A "Empty", :B "Empty"} {:_r 2, :A 1.0, :B 2.0} {:_r 3, :B 3.0})})
           (get-sheets "test/data/simple.xlsx")))
    (is (= [{:sheet []}]
           (get-sheets "missingfile.xlsx"))
        "File does not exist. Should return null.")
    (is (= [{:sheet []}] (get-sheets nil))
        "Filename was not passed in")
    (is (= [{:sheet []}] (get-sheets []))
        "Invalid argument passed in")))


(deftest get-range-test
  (testing "Get Sheet Range"
    (is (= '({:_r 1, :A "Empty", :B "Empty"} {:_r 2, :A 1.0, :B 2.0})
           (get-range (get-sheet "test/data/simple.xlsx" "Sheet1") "A1:B2")))
    (is (nil? (get-sheet-names "missingfile.xlsx"))
        "File does not exist. Should return null.")
    (is (nil? (get-sheet-names nil))
        "Filename was not passed in")))
