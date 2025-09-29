(defun c:lock-on-new-layer (/ ss i ent layer-name new-layer-name layer-obj new-layer-obj)
  (setq ss (ssget))
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq layer-name (cdr (assoc 8 (entget ent))))
        (setq new-layer-name (strcat layer-name " - lock"))

        ; Check if lock layer already exists, if not create it
        (if (not (tblsearch "LAYER" new-layer-name))
          (progn
            ; Get original layer properties
            (setq layer-obj (tblsearch "LAYER" layer-name))
            ; Create new layer with same properties as original
            (command "_-LAYER" "_N" new-layer-name "")
            ; Copy color from original layer
            (if (cdr (assoc 62 layer-obj))
              (command "_-LAYER" "_C" (cdr (assoc 62 layer-obj)) new-layer-name "")
            )
            ; Copy linetype from original layer
            (if (cdr (assoc 6 layer-obj))
              (command "_-LAYER" "_L" (cdr (assoc 6 layer-obj)) new-layer-name "")
            )
            ; Copy lineweight from original layer
            (if (cdr (assoc 370 layer-obj))
              (command "_-LAYER" "_LW" (cdr (assoc 370 layer-obj)) new-layer-name "")
            )
            ; Copy plot style from original layer
            (if (cdr (assoc 390 layer-obj))
              (command "_-LAYER" "_P" (cdr (assoc 390 layer-obj)) new-layer-name "")
            )
          )
        )

        ; Move entity to new lock layer
        (entmod (subst (cons 8 new-layer-name) (cons 8 layer-name) (entget ent)))

        (setq i (1+ i))
      )

      ; Lock all the lock layers used
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq layer-name (cdr (assoc 8 (entget ent))))
        ; Lock the layer
        (command "_-LAYER" "_LOCK" layer-name "")
        (setq i (1+ i))
      )

      (princ (strcat "\nLocked " (itoa (sslength ss)) " entities on new lock layers."))
    )
    (princ "\nNo entities selected.")
  )
  (princ)
)