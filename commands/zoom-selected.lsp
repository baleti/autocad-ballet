;; zoom-selected.lsp
;; Zooms the current view to fit around currently selected entities

(defun c:zoom-selected ( / ss i ent minpt maxpt entminpt entmaxpt)
  (setq ss (ssget "I"))  ; Get implied selection (pickfirst set)

  (if ss
    (progn
      ; Initialize min/max points
      (setq minpt nil maxpt nil)

      ; Iterate through selected entities to calculate bounding box
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (vlax-ename->vla-object (ssname ss i)))

        ; Get bounding box of current entity
        (vla-getboundingbox ent 'entminpt 'entmaxpt)

        ; Convert variants to lists
        (setq entminpt (vlax-safearray->list entminpt)
              entmaxpt (vlax-safearray->list entmaxpt))

        ; Update overall bounding box
        (if (null minpt)
          (setq minpt entminpt maxpt entmaxpt)
          (progn
            (setq minpt (list (min (car minpt) (car entminpt))
                              (min (cadr minpt) (cadr entminpt))
                              (min (caddr minpt) (caddr entminpt)))
                  maxpt (list (max (car maxpt) (car entmaxpt))
                              (max (cadr maxpt) (cadr entmaxpt))
                              (max (caddr maxpt) (caddr entmaxpt))))
          )
        )

        (setq i (1+ i))
      )

      ; Zoom to the bounding box
      (command "_.zoom" "_window" minpt maxpt)

      ; Restore the selection
      (sssetfirst nil ss)

      (princ "\nZoomed to selected objects.")
    )
    (progn
      (princ "\nNo objects selected. Select objects first, then run ZOOM-SELECTED.")
    )
  )
  (princ)
)

;; Alias for convenience
(defun c:zs () (c:zoom-selected))

(princ "\nZOOM-SELECTED command loaded. Type ZOOM-SELECTED or ZS to zoom to selected objects.")
(princ)