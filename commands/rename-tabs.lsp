(defun c:rename-tabs (/ layout-list layout-obj current-name new-name search-string replace-string count)
  ;; Generalized function to rename layout tabs by replacing specified substrings
  
  (princ "\n--- Layout Tab Renaming Tool ---")
  
  ;; Get user input for search and replace strings with paste support
  ;; The T flag allows spaces and should enable better input handling
  (setq search-string (getstring T "\nEnter string to find (Ctrl+V to paste): "))
  (if (= search-string "")
    (progn
      (princ "\nOperation cancelled - no search string entered.")
      (exit)
    )
  )
  
  (setq replace-string (getstring T "\nEnter replacement string (Ctrl+V to paste): "))
  
  (princ (strcat "\nSearching for: \"" search-string "\""))
  (princ (strcat "\nReplacing with: \"" replace-string "\""))
  (princ "\n")
  
  ;; Initialize counter
  (setq count 0)
  
  ;; Get all layout objects
  (setq layout-list (layoutlist))
  
  ;; Loop through each layout
  (foreach layout-name layout-list
    (if (/= layout-name "Model") ; Skip Model space
      (progn
        (setq layout-obj (vla-item (vla-get-layouts (vla-get-activedocument (vlax-get-acad-object))) layout-name))
        (setq current-name (vla-get-name layout-obj))
        
        ;; Check if the layout name contains the search string
        (if (vl-string-search search-string current-name)
          (progn
            ;; Replace search string with replacement string
            (setq new-name (vl-string-subst replace-string search-string current-name))
            
            ;; Only rename if the new name is different
            (if (/= current-name new-name)
              (progn
                (vla-put-name layout-obj new-name)
                (setq count (+ count 1))
                (princ (strcat "\nRenamed: \"" current-name "\" -> \"" new-name "\""))
              )
            )
          )
        )
      )
    )
  )
  
  (princ (strcat "\n\nOperation complete! " (itoa count) " tab(s) renamed."))
  (princ)
)