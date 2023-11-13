(ns core-test
  (:require [clojure.java.io :as io]
            [clojure.test :refer [deftest is run-tests testing]]
            [bb-excel.core :refer [get-sheets get-sheet-names get-sheet
                                   get-range
                                   ; get-row get-col get-cells crange
                                   ]])
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

(comment
  (run-tests)

  (->>
   (get-sheets "test/data/Types.xlsx")
   second
   :sheet
   clojure.pprint/print-table))
