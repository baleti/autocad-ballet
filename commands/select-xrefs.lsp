(defun c:select-xrefs (/ ss i ent xref-list)
  ;; Initialize variables
  (setq xref-list '())
  (setq ss (ssget "_X" '((0 . "INSERT"))))
  
  ;; Check if any INSERT entities were found
  (if ss
    (progn
      ;; Loop through all INSERT entities
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        ;; Check if the INSERT is an xref
        (if (= (logand 4 (cdr (assoc 70 (tblsearch "BLOCK" 
                                                   (cdr (assoc 2 (entget ent)))))))
               4)
          ;; Add to xref list if it's an xref
          (setq xref-list (cons ent xref-list))
        )
        (setq i (1+ i))
      )
      
      ;; Create selection set with xrefs only
      (if xref-list
        (progn
          ;; Clear any previous selection
          (sssetfirst nil nil)
          ;; Create new selection set
          (setq ss (ssadd))
          (foreach xref xref-list
            (ssadd xref ss)
          )
          ;; Select the xrefs
          (sssetfirst nil ss)
          (princ (strcat "\n" (itoa (length xref-list)) " xref(s) selected."))
        )
        (princ "\nNo xrefs found in the drawing.")
      )
    )
    (princ "\nNo block references found in the drawing.")
  )
  (princ)
)