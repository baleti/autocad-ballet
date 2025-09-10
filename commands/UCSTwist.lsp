(defun c:ucs-twist (/ listbox lst ucs)
  
  (defun listbox (title keys / tmp file dcl_id choice)
  (setq	tmp  (vl-filename-mktemp "tmp.dcl")
	file (open tmp "w")
  )
  (write-line
    (strcat "ListBox:dialog{label=\"" title "\";")
    file
  )
  (write-line "spacer;:list_box{key=\"lst\";}spacer;ok_only;}" file)
  (close file)
  (setq dcl_id (load_dialog tmp))
  (if (not (new_dialog "ListBox" dcl_id))
    (exit)
  )
  (start_list "lst")
  (mapcar 'add_list keys)
  (end_list)
  (action_tile "lst" "(setq choice (nth (atoi $value) keys)) (done_dialog)")
  (start_dialog)
  (unload_dialog dcl_id)
  (vl-file-delete tmp)
  choice
)
  
  (while (setq ucs (tblnext "UCS" (null ucs)))
    (setq lst (cons (cdr (assoc 2 ucs)) lst))
  )
  (setq lst (cons "World" (cons "Previous" (reverse lst))))
  (if (setq ucs (listbox "UCS Twist" lst))
    (progn
      (cond
	((= ucs "World") (command-s "_.ucs" "_world"))
	((= ucs "Previous") (command-s "_.ucs" "_previous"))
	(T (command-s "_.ucs" "_restore" ucs))
      )
      (command-s "_.dview"
		 ""
		 "_twist"
		 (angtos
		   (angle '(0 0) (trans '(1 0) 0 1 T))
		   (getvar 'aunits)
		   16
		 )
		 ""
      )
    )
  )
  (princ)
)