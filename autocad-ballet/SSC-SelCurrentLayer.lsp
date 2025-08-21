(defun c:ssc ()
   (sssetfirst nil (ssget "X" (list (cons 8 (getvar "clayer")))))
   (princ)
)