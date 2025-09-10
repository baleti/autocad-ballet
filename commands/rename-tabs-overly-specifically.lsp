(defun c:rename-tabs-overly-specifically (/ layout-list layout-name layout-obj current-name 
                                          new-name drawing-num ss title-block att-list 
                                          att count code-end-pos)
  ;; The most ridiculously specific tab renaming command ever created
  ;; Because why be flexible when you can hardcode everything?
  
  (princ "\n--- OVERLY SPECIFIC Tab Renaming Tool ---")
  (princ "\n(For BT-2602-BM3-BM-XXX format ONLY, using BM3 Title Block A3 ONLY)")
  (princ "\n")
  
  ;; Initialize counter
  (setq count 0)
  
  ;; Get all layout objects
  (setq layout-list (layoutlist))
  
  ;; Loop through each layout
  (foreach layout-name layout-list
    (if (/= layout-name "Model") ; Skip Model space (as explicitly requested!)
      (progn
        ;; Set the layout as current to search for blocks
        (setvar "CTAB" layout-name)
        
        ;; Get current layout object
        (setq layout-obj (vla-item 
                          (vla-get-layouts 
                            (vla-get-activedocument 
                              (vlax-get-acad-object))) 
                          layout-name))
        (setq current-name (vla-get-name layout-obj))
        
        ;; Check if this tab name matches our VERY SPECIFIC format
        ;; Looking for pattern like "BT-2602-BM3-BM-XXX - DESCRIPTION"
        (if (and 
              (vl-string-search " - " current-name) ; Must have the separator
              (vl-string-search "BT-" current-name) ; Must start with BT-
              (= (vl-string-search "BT-" current-name) 0)) ; BT- must be at the beginning
          (progn
            ;; Find the position of " - " to extract the code part
            (setq code-end-pos (vl-string-search " - " current-name))
            
            ;; Search for our VERY SPECIFIC block: "BM3 Title Block A3"
            (setq ss (ssget "X" 
                           (list 
                             '(0 . "INSERT") 
                             '(2 . "BM3 Title Block A3")
                             (cons 410 layout-name))))
            
            (if ss
              (progn
                ;; Get the first (and hopefully only) title block
                (setq title-block (vlax-ename->vla-object (ssname ss 0)))
                
                ;; Initialize drawing number as nil
                (setq drawing-num nil)
                
                ;; Get attributes
                (if (= (vla-get-hasattributes title-block) :vlax-true)
                  (progn
                    (setq att-list (vlax-safearray->list 
                                    (vlax-variant-value 
                                      (vla-getattributes title-block))))
                    
                    ;; Find the DRAWING_NUMBER attribute (case sensitive because we're THAT specific)
                    (foreach att att-list
                      (if (= (vla-get-tagstring att) "DRAWING_NUMBER")
                        (setq drawing-num (vla-get-textstring att))
                      )
                    )
                  )
                )
                
                ;; If we found a drawing number, do the replacement
                (if drawing-num
                  (progn
                    ;; Extract everything after " - " from the original name
                    (setq new-name (strcat 
                                    drawing-num 
                                    (substr current-name 
                                           (+ code-end-pos 1) ; Start from " - "
                                           (strlen current-name))))
                    
                    ;; Only rename if the new name is different
                    (if (/= current-name new-name)
                      (progn
                        (vla-put-name layout-obj new-name)
                        (setq count (+ count 1))
                        (princ (strcat "\nRenamed: \"" current-name "\""))
                        (princ (strcat "\n      to: \"" new-name "\""))
                      )
                      (princ (strcat "\nSkipped (already correct): \"" current-name "\""))
                    )
                  )
                  (princ (strcat "\nWarning: No DRAWING_NUMBER attribute found in block on tab: " layout-name))
                )
              )
              (princ (strcat "\nWarning: No 'BM3 Title Block A3' found on tab: " layout-name))
            )
          )
          (princ (strcat "\nSkipped (wrong format): \"" current-name "\""))
        )
      )
    )
  )
  
  (princ "\n")
  (princ "\n===========================================")
  (princ (strcat "\nOperation complete! " (itoa count) " tab(s) renamed."))
  (princ "\nThis command was brought to you by:")
  (princ "\n  - The hardcoded value 'BM3 Title Block A3'")
  (princ "\n  - The assumption that all tabs start with 'BT-'")
  (princ "\n  - The belief that ' - ' is the universal separator")
  (princ "\n===========================================")
  (princ)
)