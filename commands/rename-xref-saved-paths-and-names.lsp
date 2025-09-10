(defun c:rename-xref-saved-paths-and-names (/ block-defs block-def xref-path new-path xref-name new-name
                                         search-string replace-string count-paths count-names doc)
  ;; Function to rename both saved paths AND reference names in xrefs by replacing specified substrings
  
  (princ "\n--- Xref Saved Path and Reference Name Renaming Tool ---")
  
  ;; Get user input for search and replace strings with paste support
  ;; The T flag allows spaces and should enable better input handling
  (setq search-string (getstring T "\nEnter string to find in xref paths and names (Ctrl+V to paste): "))
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
  
  ;; Initialize counters
  (setq count-paths 0)
  (setq count-names 0)
  
  ;; Get the active document
  (setq doc (vla-get-activedocument (vlax-get-acad-object)))
  
  ;; Get the blocks collection
  (setq block-defs (vla-get-blocks doc))
  
  ;; Loop through each block definition
  (vlax-for block-def block-defs
    ;; Check if the block is an xref
    (if (= (vla-get-isxref block-def) :vlax-true)
      (progn
        ;; PART 1: Handle the saved path
        (setq xref-path (vla-get-path block-def))
        
        ;; Check if the path contains the search string
        (if (vl-string-search search-string xref-path)
          (progn
            ;; Replace search string with replacement string in path
            (setq new-path (vl-string-subst replace-string search-string xref-path))
            
            ;; Only update if the new path is different
            (if (/= xref-path new-path)
              (progn
                ;; Attempt to set the new path
                (if (not (vl-catch-all-error-p 
                          (vl-catch-all-apply 'vla-put-path (list block-def new-path))))
                  (progn
                    (setq count-paths (+ count-paths 1))
                    (princ (strcat "\nRenamed PATH for \"" (vla-get-name block-def) "\": "))
                    (princ (strcat "\n  From: \"" xref-path "\""))
                    (princ (strcat "\n  To:   \"" new-path "\""))
                  )
                  (princ (strcat "\n  Warning: Could not update path for \"" 
                               (vla-get-name block-def) 
                               "\" - new path may be invalid"))
                )
              )
            )
          )
        )
        
        ;; PART 2: Handle the reference name
        (setq xref-name (vla-get-name block-def))
        
        ;; Check if the name contains the search string
        (if (vl-string-search search-string xref-name)
          (progn
            ;; Replace search string with replacement string in name
            (setq new-name (vl-string-subst replace-string search-string xref-name))
            
            ;; Only update if the new name is different
            (if (/= xref-name new-name)
              (progn
                ;; Attempt to set the new name
                (if (not (vl-catch-all-error-p 
                          (vl-catch-all-apply 'vla-put-name (list block-def new-name))))
                  (progn
                    (setq count-names (+ count-names 1))
                    (princ (strcat "\nRenamed NAME: \"" xref-name "\" -> \"" new-name "\""))
                  )
                  (princ (strcat "\n  Warning: Could not rename \"" 
                               xref-name 
                               "\" to \"" 
                               new-name
                               "\" - name may already exist or be invalid"))
                )
              )
            )
          )
        )
      )
    )
  )
  
  ;; Reload xrefs to reflect changes
  (if (or (> count-paths 0) (> count-names 0))
    (progn
      (princ "\n\nReloading xrefs...")
      (vl-cmdf "_.XREF" "_RELOAD" "*")
    )
  )
  
  (princ (strcat "\n\nOperation complete!"))
  (princ (strcat "\n  " (itoa count-paths) " xref path(s) renamed"))
  (princ (strcat "\n  " (itoa count-names) " xref name(s) renamed"))
  (princ)
)

;; Advanced version with additional features
(defun c:rename-xref-saved-paths-and-names-advanced (/ block-defs block-def xref-path new-path xref-name new-name
                                                search-string replace-string count-paths count-names doc 
                                                case-sensitive xref-list choice rename-choice)
  ;; Advanced version with case sensitivity option and selective renaming
  
  (princ "\n--- Advanced Xref Path and Name Renaming Tool ---")
  
  ;; Get the active document
  (setq doc (vla-get-activedocument (vlax-get-acad-object)))
  
  ;; First, list all current xrefs and their paths
  (princ "\n\nCurrent Xrefs in drawing:")
  (princ "\n" )
  (setq block-defs (vla-get-blocks doc))
  (setq xref-list nil)
  
  (vlax-for block-def block-defs
    (if (= (vla-get-isxref block-def) :vlax-true)
      (progn
        (princ (strcat "\n  Name: \"" (vla-get-name block-def) "\""))
        (princ (strcat "\n  Path: \"" (vla-get-path block-def) "\""))
        (princ "\n")
        (setq xref-list (cons (vla-get-name block-def) xref-list))
      )
    )
  )
  
  (if (null xref-list)
    (progn
      (princ "\n\nNo xrefs found in drawing.")
      (exit)
    )
  )
  
  ;; Ask what to rename
  (initget "Paths Names Both")
  (setq rename-choice (getkword "\nRename [Paths/Names/Both] <Both>: "))
  (if (null rename-choice) (setq rename-choice "Both"))
  
  ;; Ask for case sensitivity
  (initget "Yes No")
  (setq choice (getkword "\nCase sensitive search? [Yes/No] <No>: "))
  (setq case-sensitive (= choice "Yes"))
  
  ;; Get user input for search and replace strings
  (setq search-string (getstring T "\nEnter string to find (Ctrl+V to paste): "))
  (if (= search-string "")
    (progn
      (princ "\nOperation cancelled - no search string entered.")
      (exit)
    )
  )
  
  (setq replace-string (getstring T "\nEnter replacement string (Ctrl+V to paste): "))
  
  (princ (strcat "\n\nSearching for: \"" search-string "\""))
  (princ (strcat "\nReplacing with: \"" replace-string "\""))
  (princ (strcat "\nRenaming: " rename-choice))
  (if case-sensitive
    (princ " (case sensitive)")
    (princ " (case insensitive)")
  )
  (princ "\n")
  
  ;; Initialize counters
  (setq count-paths 0)
  (setq count-names 0)
  
  ;; Loop through each block definition
  (vlax-for block-def block-defs
    ;; Check if the block is an xref
    (if (= (vla-get-isxref block-def) :vlax-true)
      (progn
        ;; PART 1: Handle the saved path (if requested)
        (if (or (= rename-choice "Paths") (= rename-choice "Both"))
          (progn
            (setq xref-path (vla-get-path block-def))
            
            ;; Perform search based on case sensitivity setting
            (if case-sensitive
              ;; Case sensitive search
              (if (vl-string-search search-string xref-path)
                (setq new-path (vl-string-subst replace-string search-string xref-path))
                (setq new-path nil)
              )
              ;; Case insensitive search
              (if (vl-string-search (strcase search-string) (strcase xref-path))
                (setq new-path (ReplaceStringCaseInsensitive xref-path search-string replace-string))
                (setq new-path nil)
              )
            )
            
            ;; Only update if a match was found and the new path is different
            (if (and new-path (/= xref-path new-path))
              (progn
                ;; Attempt to set the new path
                (if (not (vl-catch-all-error-p 
                          (vl-catch-all-apply 'vla-put-path (list block-def new-path))))
                  (progn
                    (setq count-paths (+ count-paths 1))
                    (princ (strcat "\nRenamed PATH for \"" (vla-get-name block-def) "\": "))
                    (princ (strcat "\n  From: \"" xref-path "\""))
                    (princ (strcat "\n  To:   \"" new-path "\""))
                  )
                  (princ (strcat "\n  Warning: Could not update path for \"" 
                               (vla-get-name block-def) 
                               "\" - new path may be invalid"))
                )
              )
            )
          )
        )
        
        ;; PART 2: Handle the reference name (if requested)
        (if (or (= rename-choice "Names") (= rename-choice "Both"))
          (progn
            (setq xref-name (vla-get-name block-def))
            
            ;; Perform search based on case sensitivity setting
            (if case-sensitive
              ;; Case sensitive search
              (if (vl-string-search search-string xref-name)
                (setq new-name (vl-string-subst replace-string search-string xref-name))
                (setq new-name nil)
              )
              ;; Case insensitive search
              (if (vl-string-search (strcase search-string) (strcase xref-name))
                (setq new-name (ReplaceStringCaseInsensitive xref-name search-string replace-string))
                (setq new-name nil)
              )
            )
            
            ;; Only update if a match was found and the new name is different
            (if (and new-name (/= xref-name new-name))
              (progn
                ;; Attempt to set the new name
                (if (not (vl-catch-all-error-p 
                          (vl-catch-all-apply 'vla-put-name (list block-def new-name))))
                  (progn
                    (setq count-names (+ count-names 1))
                    (princ (strcat "\nRenamed NAME: \"" xref-name "\" -> \"" new-name "\""))
                  )
                  (princ (strcat "\n  Warning: Could not rename \"" 
                               xref-name 
                               "\" to \"" 
                               new-name
                               "\" - name may already exist or be invalid"))
                )
              )
            )
          )
        )
      )
    )
  )
  
  ;; Reload xrefs to reflect changes
  (if (or (> count-paths 0) (> count-names 0))
    (progn
      (princ "\n\nReloading xrefs...")
      (vl-cmdf "_.XREF" "_RELOAD" "*")
    )
  )
  
  (princ (strcat "\n\nOperation complete!"))
  (princ (strcat "\n  " (itoa count-paths) " xref path(s) renamed"))
  (princ (strcat "\n  " (itoa count-names) " xref name(s) renamed"))
  (princ)
)

;; Helper function for case-insensitive string replacement
(defun ReplaceStringCaseInsensitive (str search replace / pos len-search result)
  ;; Replace search string with replace string case-insensitively
  (setq len-search (strlen search))
  (setq result "")
  
  (while (setq pos (vl-string-search (strcase search) (strcase str)))
    (setq result (strcat result (substr str 1 pos) replace))
    (setq str (substr str (+ pos len-search 1)))
  )
  (strcat result str)
)