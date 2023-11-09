(ns bb-excel.util)

(defn by-tag
  "Pred to use in fns like filter to choose xml nodes by tag"
  [value]
  (fn [node]
    (= value (:tag node))))

(defn find-first
  "Finds the first item in a collection that matches a predicate. Returns a
  transducer when no collection is provided.
  Taken from https://github.com/weavejester/medley"
  ([pred]
   (fn [rf]
     (fn
       ([] (rf))
       ([result] (rf result))
       ([result x]
        (if (pred x)
          (ensure-reduced (rf result x))
          result)))))
  ([pred coll]
   (reduce (fn [_ x] (when (pred x) (reduced x))) nil coll)))

(defn throw-ex
  [message]
  (throw (ex-info message {})))
