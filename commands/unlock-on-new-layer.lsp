(defun c:unlock-on-new-layer (/ ss i ent layer-name orig-layer-name layers-to-check 
                               processed-count lock-layers pt1 pt2 temp-ss all-ents
                               selected-ents ent-list)
  (princ "\nSelect entities to unlock (use Window/Crossing for locked entities): ")
  
  ; First, get all lock layers
  (setq lock-layers '())
  (setq layer-name (tblnext "LAYER" T))
  (while layer-name
    (setq layer-name (cdr (assoc 2 layer-name)))
    (if (and (> (strlen layer-name) 7)
             (= (substr layer-name (- (strlen layer-name) 6)) " - lock"))
      (setq lock-layers (cons layer-name lock-layers))
    )
    (setq layer-name (tblnext "LAYER"))
  )
  
  ; Temporarily unlock all lock layers to allow selection
  (foreach lock-layer lock-layers
    (command "_-LAYER" "_UNLOCK" lock-layer "")
  )
  
  ; Now get selection (will work on previously locked layers)
  (setq ss (ssget))
  
  ; Process selection
  (if ss
    (progn
      (setq processed-count 0)
      (setq layers-to-check '())
      
      ; Process each entity in selection
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i))
        (setq layer-name (cdr (assoc 8 (entget ent))))
        
        ; Check if this entity is on a lock layer
        (if (and (> (strlen layer-name) 7)
                 (= (substr layer-name (- (strlen layer-name) 6)) " - lock"))
          (progn
            ; Extract original layer name
            (setq orig-layer-name (substr layer-name 1 (- (strlen layer-name) 7)))
            
            ; Check if original layer exists
            (if (tblsearch "LAYER" orig-layer-name)
              (progn
                ; Move entity back to original layer
                (entmod (subst (cons 8 orig-layer-name) 
                              (cons 8 layer-name) 
                              (entget ent)))
                (setq processed-count (1+ processed-count))
                
                ; Add lock layer to list for cleanup check
                (if (not (member layer-name layers-to-check))
                  (setq layers-to-check (cons layer-name layers-to-check))
                )
              )
              (princ (strcat "\nWarning: Original layer '" orig-layer-name 
                            "' not found for entity on layer '" layer-name "'"))
            )
          )
        )
        (setq i (1+ i))
      )
      
      ; Re-lock any lock layers that still have entities
      (foreach lock-layer lock-layers
        ; Check if any entities exist on this layer
        (setq temp-ss (ssget "X" (list (cons 8 lock-layer))))
        (if temp-ss
          ; Layer has entities, lock it again
          (command "_-LAYER" "_LOCK" lock-layer "")
          ; Layer is empty, delete it
          (progn
            (command "_-LAYER" "_DELETE" lock-layer "")
            (princ (strcat "\nDeleted empty lock layer: " lock-layer))
          )
        )
      )
      
      ; Report results
      (if (> processed-count 0)
        (princ (strcat "\nUnlocked " (itoa processed-count) 
                      " entities and moved to original layers."))
        (princ "\nNo entities on lock layers found in selection.")
      )
    )
    (princ "\nNo entities selected.")
  )
  (princ)
)