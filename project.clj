(defproject com.github.kbosompem/bb-excel "0.0.4"
  :description "A Simple Clojure/Babashka Library for Reading Data from Excel Files"
  :url "https://github.com/kbosompem/bb-excel"
  :license {:name "EPL-2.0"
            :url "https://www.eclipse.org/legal/epl-2.0/"}
  :dependencies [[org.clojure/clojure "1.11.1"]
                 [org.clojure/data.xml "0.2.0-alpha6"]
                 [org.clojure/tools.cli "1.0.206"]]
  :plugins [[lein-codox "0.10.8"]
            [lein-ancient "1.0.0-RC3"]]
  :deploy-repositories [["releases" :clojars]
                        ["snapshots" :clojars]]
  :main bb-excel.core
  :uberjar-name "bb-excel.jar"
  :jar-name "bb-excel-slim.jar"
  :aot :all
  :repl-options {:init-ns bb-excel.core})