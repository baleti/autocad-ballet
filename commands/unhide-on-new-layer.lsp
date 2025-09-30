(defun c:unhide-on-new-layer (/ i ent layer-name orig-layer-name
                               processed-count hide-layers temp-ss)
  ; Get all hide layers
  (setq hide-layers '())
  (setq layer-name (tblnext "LAYER" T))
  (while layer-name
    (setq layer-name (cdr (assoc 2 layer-name)))
    (if (and (> (strlen layer-name) 7)
             (= (substr layer-name (- (strlen layer-name) 6)) " - hide"))
      (setq hide-layers (cons layer-name hide-layers))
    )
    (setq layer-name (tblnext "LAYER"))
  )

  (if hide-layers
    (progn
      (setq processed-count 0)

      ; Process each hide layer
      (foreach hide-layer hide-layers
        ; Thaw the layer to access entities
        (command "_-LAYER" "_THAW" hide-layer "")

        ; Get all entities on this hide layer
        (setq temp-ss (ssget "X" (list (cons 8 hide-layer))))

        (if temp-ss
          (progn
            ; Extract original layer name
            (setq orig-layer-name (substr hide-layer 1 (- (strlen hide-layer) 7)))

            ; Check if original layer exists
            (if (tblsearch "LAYER" orig-layer-name)
              (progn
                ; Move all entities back to original layer
                (setq i 0)
                (repeat (sslength temp-ss)
                  (setq ent (ssname temp-ss i))
                  (entmod (subst (cons 8 orig-layer-name)
                                (cons 8 hide-layer)
                                (entget ent)))
                  (setq processed-count (1+ processed-count))
                  (setq i (1+ i))
                )

                ; Delete the now-empty hide layer
                (command "_-LAYER" "_DELETE" hide-layer "")
                (princ (strcat "\nDeleted hide layer: " hide-layer))
              )
              (progn
                ; Original layer doesn't exist, keep hide layer but freeze it
                (princ (strcat "\nWarning: Original layer '" orig-layer-name
                              "' not found. Keeping hide layer '" hide-layer "' frozen."))
                (command "_-LAYER" "_FREEZE" hide-layer "")
              )
            )
          )
          ; Layer is already empty, delete it
          (progn
            (command "_-LAYER" "_DELETE" hide-layer "")
            (princ (strcat "\nDeleted empty hide layer: " hide-layer))
          )
        )
      )

      ; Report results
      (if (> processed-count 0)
        (princ (strcat "\nUnhidden " (itoa processed-count)
                      " entities and moved to original layers."))
        (princ "\nNo entities found on hide layers.")
      )
    )
    (princ "\nNo hide layers found.")
  )
  (princ)
)