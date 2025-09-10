(defun c:select-current-layer ()
   (sssetfirst nil (ssget "X" (list (cons 8 (getvar "clayer")))))
   (princ)
)
