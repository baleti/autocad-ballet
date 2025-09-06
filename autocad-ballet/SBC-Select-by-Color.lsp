(defun c:sbc ( / c d e l )
   (if (setq e (car (entsel)))
       (progn
           (setq c
               (cond
                   (   (cdr (assoc 62 (entget e)))   )
                   (   (abs (cdr (assoc 62 (tblsearch "LAYER" (cdr (assoc 8 (entget 

e)))))))   )
               )
           )                     
           (while (setq d (tblnext "LAYER" (null d)))
               (if (= c (abs (cdr (assoc 62 d))))
                   (setq l (cons "," (cons (cdr (assoc 2 d)) l)))
               )
           )
           (sssetfirst nil
               (ssget "_X"
                   (if l
                       (list
                           (cons -4 "<OR")
                               (cons 62 c)
                               (cons -4 "<AND")
                                   (cons 62 256)
                                   (cons 8 (apply 'strcat (cdr l)))
                               (cons -4 "AND>")
                           (cons -4 "OR>")
                       )
                       (list (cons 62 c))
                   )
               )
           )
       )
   )
   (princ)
)