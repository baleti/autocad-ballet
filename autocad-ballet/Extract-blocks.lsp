;;Make individual dwgs from all the blocks in a drawing
;; By Mark S. Thomas

(defun c:bwblock (/ ss cntr bn)
 (setq ss (ssget "x" '((0 . "INSERT")))
       cntr 0
        )

 (mapcar 'setvar '(filedia cmddia) '(0 0))

 (if (not (zerop (sslength ss)))
   (while (setq ent (ssname ss cntr))
      (setq bn (cdr (assoc 2 (entget ent)))
            cntr (1+ cntr)
            )
     ;(command "_wblock" (strcat "G:\\0000 CAD LIBRARY\Blocks\\" bn) bn "NO"); for Map
     (command "_wblock" (strcat "G:\\0000 CAD LIBRARY\\Blocks\\" bn) bn "NO")
     )
   )
 (mapcar 'setvar '(filedia cmddia) '(1 1))
 )