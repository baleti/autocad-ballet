;;;  force save all open drawings without prompts
;;;  silently skip read-only drawings and drawings without a location
(defun C:SAVE-ALL-FORCE (/ *error* echo docs doc)
  
  (defun *error* (msg)
    (if echo (setvar "CMDECHO" echo))
    (if (and msg 
             (not (member msg '("Function cancelled" "quit / exit abort"))))
      (princ (strcat "\nError: " msg)))
    (princ))
  
  (vl-load-com)
  (setq echo (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  
  (setq docs (vla-get-documents (vlax-get-acad-object)))
  
  (vlax-for doc docs
    (if (and
          (= 1 (vlax-variant-value (vla-getvariable doc "DWGTITLED")))
          (= 1 (vlax-variant-value (vla-getvariable doc "WRITESTAT")))
          (= "" (vlax-variant-value (vla-getvariable doc "REFEDITNAME"))))
      (vl-catch-all-apply 'vla-save (list doc))))
  
  (setvar "CMDECHO" echo)
  (princ "\nAll saveable drawings processed.")
  (princ))

(princ "\nSAVE-ALL-FORCE loaded. Type SAVE-ALL-FORCE to run.")
(princ)
