# bb-excel
[![Clojars Project](https://img.shields.io/clojars/v/com.github.kbosompem/bb-excel.svg)](https://clojars.org/com.github.kbosompem/bb-excel)
[![bb compatible](https://raw.githubusercontent.com/babashka/babashka/master/logo/badge.svg)](https://babashka.org)

Use [**Babashka**](https://www.babashka.org) to extract data from Excel Spreadsheets!

bb-excel is a simple [**Clojure**](https://www.clojure.org) library to extract data from Excel files without relying on the Apache POI OOXML library. See below for rationale.

## Installation

Add the following dependency to your `project.clj` , `deps.edn`, or `bb.edn`

[![Clojars Project](https://clojars.org/com.github.kbosompem/bb-excel/latest-version.svg)](https://clojars.org/com.github.kbosompem/bb-excel)

## BBIN install
You can also install it as a [**bbin**](https://github.com/babashka/bbin) utility to convert Excel Sheets to EDN
```bash
bbin install com.github.kbosompem/bb-excel --latest-sha
```


## Status

Please note that being a beta version indicates the provisional status
of this library, and that features are subject to change.


## Getting started


This library is meant for the simplest of spreadsheets.

The primary function is **get-sheet** 

In the first example below we pull data from a specific sheet within the workbook. It returns a list of hash-maps with :_r key indicating the row within the sheet. 

```clojure
#!/usr/bin/env bb
(require '[babashka.deps :as deps])
(deps/add-deps 
  '{:deps {com.github.kbosompem/bb-excel {:mvn/version "0.0.5"}}}) 

(ns demo
  (:require [clojure.pprint :refer [print-table pprint]]
            [bb-excel.core  :refer [get-sheet get-sheets]]))

(get-sheet "test/data/simple.xlsx" "Shows" )

;=>
({:_r 2, :A 1.0, :B "Sesame Street"}
 {:_r 3, :A 2.0, :B "La Femme Nikita"}
 {:_r 4, :A 3.0, :B "House M.D"}
 {:_r 5, :A 4.0, :B "Breaking Bad"})

```

Let's print out the results for better readability

```clojure
(print-table
   (get-sheet "test/data/simple.xlsx" "Shows" ))

| :_r |   :A |              :B |
|-----+------+-----------------|
|   1 | Rank |         TV Show |
|   2 |  1.0 |   Sesame Street |
|   3 |  2.0 | La Femme Nikita |
|   4 |  3.0 |       House M.D |
|   5 |  4.0 |    Breaking Bad |
```

If there is a header row we can use the values from that row as keys.Set :hdr flag in opts to true and specify the row number.

```clojure
(print-table
   (get-sheet "test/data/simple.xlsx" "Shows" {:hdr true :row 1}))

| :_r | Rank |         TV Show |
|-----+------+-----------------|
|   2 |  1.0 |   Sesame Street |
|   3 |  2.0 | La Femme Nikita |
|   4 |  3.0 |       House M.D |
|   5 |  4.0 |    Breaking Bad |
```
You can also format the keys by providing a function. In the example below we keywordize each heading

```clojure
(print-table
   (get-sheet "test/data/simple.xlsx" "Shows" {:hdr true :fxn keyword :row 1}))

| :_r | :Rank |        :TV Show |
|-----+-------+-----------------|
|   2 |   1.0 |   Sesame Street |
|   3 |   2.0 | La Femme Nikita |
|   4 |   3.0 |       House M.D |
|   5 |   4.0 |    Breaking Bad | 
```

Another example. skeyword Lowercases, replaces spaces and slashes, excises non-ascii characters and limits to 50 characters every row heading

```clojure
(print-table
   (get-sheet "test/data/simple.xlsx" "Shows" 
              {:hdr true 
               :fxn bb-excel.cli/skeyword}))

| :_r | :rank |        :tv_show |
|-----+-------+-----------------|
|   2 |   1.0 |   Sesame Street |
|   3 |   2.0 | La Femme Nikita |
|   4 |   3.0 |       House M.D |
|   5 |   4.0 |    Breaking Bad |
```
That's all well and good but how do we use the same cell addressing format used in excel?
For instance if we want B2:C4 then we can leverage the get-range function instead.
Now that expects a sheet so 

```clojure
(print-table
    (get-range
     (get-sheet "test/data/simple.xlsx" "Shows")
     "A3:B4"))

| :_r |  :A |              :B |
|-----+-----+-----------------|
|   3 | 2.0 | La Femme Nikita |
|   4 | 3.0 |       House M.D |
```

## Limitations

1. Does not support older xls format
2. No write planned at this moment


## Rationale

Why create another excel library in clojure when you can use docjure or wrap the venerable Apache POI library.
Answer is simple. [**Babashka**](https://www.babashka.org), my scripting language of choice does not support Apache POI.
This library currently reads xlsx files and returns a vector of hashmaps. Each hashmap representing a row in a selected sheet.



## Recommended Alternatives

1. If you are using clojure and need to read from and write to Excel files I highly recommend you take a look at Martin Jul's  [**DOCJURE**](https://github.com/mjul/docjure)
2. If you need to generate excel files take a look at Mathew Downey's [**EXCEL-CLJ**](https://github.com/matthewdowney/excel-clj)
3. If you need to generate multi-format reports try [JasperReport](https://sourceforge.net/projects/jasperreports/)
