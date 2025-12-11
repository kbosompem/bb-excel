(defproject com.github.kbosompem/bb-excel "0.2.0"
  :description "A Simple Babashka Library for working with Microsoft Excel Files"
  :url "https://github.com/kbosompem/bb-excel"
  :license {:name "EPL-2.0"
            :url "https://www.eclipse.org/legal/epl-2.0/"}
  :dependencies [[org.clojure/clojure "1.11.1"]
                 [org.clojure/data.xml "0.2.0-alpha6"]
                 [hiccup "2.0.0-RC2"]
                 [metosin/malli  "0.9.0"]
                 [org.clojure/tools.cli "1.0.219"]]
  :plugins [[lein-ancient "1.0.0-RC3"]]
  :deploy-repositories [["releases" {:url "https://repo.clojars.org"
                                     :username :env/clojars_username
                                     :password :env/clojars_password
                                     :sign-releases false}]
                        ["snapshots" {:url "https://repo.clojars.org"
                                      :username :env/clojars_username
                                      :password :env/clojars_password}]]
  :main bb-excel.core
  :uberjar-name "bb-excel.jar"
  :jar-name "bb-excel-slim.jar"
  :aot :all
  :repl-options {:init-ns bb-excel.core})
