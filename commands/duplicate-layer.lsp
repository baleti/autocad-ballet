;;; ================================================
;                     DUPLICATE_LAYER
;;; ================================================
;Author: Andrzej Kalinowski
;26.08.2016
;EN: Lisp creates new layer with the same properties as layer of selected object
;PL: lisp tworzy nowa taka sama warstwe jak wskazany obiekt i pyta o nazwe

(defun c:DD (/ obj1 laylist Lname1 uInput entlist1 NrNotSpaces)
    ;-----------------------------------------
    ;selection of base object
    ;-----------------------------------------
    (setq obj1 nil)
    (prompt "\nSelect object ")
    (while (= obj1 nil)
        (setq obj1 (ssget "_+.:S") )
        (if (= obj1 nil)
            (princ "\nwrong object");then
            (if (> (sslength obj1) 1);else-spr czy wybrano tylko 1 obj
                (progn ;gdy wybrano za duzo obiektow
                    (princ "\nselect only one object") (princ)
                    (setq obj1 nil)
                );progn
                (progn ;else
                    (setq
                        obj1 (ssname obj1 0)
                        entlist1 (entget obj1)
                        Lname1 (cdr (assoc 8 entlist1) )
                    );end setq
                );progn else
            );end if
        );end if
    );end while
    
    ;-----------------------------------------
    ;getting new name
    ;-----------------------------------------
    ;(setq Bname2 "sth1")
    (setq uInput (SETNEWNAME_DCL Lname1 "OFF") )
    (setq NrNotSpaces (CHECKSPACES uInput ) )
        ;SETNEWNAME_DCL function opens DCL window i writes old name in the active field in case you want to change only part of it
    (while 
        (or 
            (= uInput nil)
            (= uInput "")
            ;(= uInput Lname1)
            (= NrNotSpaces 0);checks if in uInput there are other signs then " "
            (and (/= uInput Lname1) (/=  (tblsearch "layer" uInput) nil ) )
            ;(/=  (tblsearch "Block" uInput) nil )
        );or  

        (cond
            (;condition1
                (or (= NrNotSpaces 0) (= nil uInput) (= "" uInput) )
                (setq
                    uInput (SETNEWNAME_DCL Lname1 "OFF")
                    NrNotSpaces (CHECKSPACES uInput )
                );setq
            );condition1
            (;condition2
                (/=  (tblsearch "layer" uInput) nil )
                (setq
                    uInput (SETNEWNAME_DCL Lname1 "ON")
                    NrNotSpaces (CHECKSPACES uInput )
                );setq
            );condition2
        );cond
    );while
    
    ;-----------------------------------------
    ;creating new layer and setting its properties
    ;-----------------------------------------
    (setq
        laylist (entget (tblobjname "layer" Lname1))
    	laylist (subst (cons 2 uInput) (assoc 2 laylist) laylist)
    );setq
    (entmake laylist)
    (command "_-layer" "_set" uInput "")
     
    (princ "\t\tOld layername:   ")(princ uInput)
    (princ)
);end defun
;;; ================================================
;                      CHECKSPACES
;;; ================================================
(defun CHECKSPACES ( NameToCheck1 / count1 result2 )
    (setq
        result2 0
        count1 1
    );setq
    (repeat (strlen NameToCheck1)
        (if (/= (substr NameToCheck1 count1 1) " ")
            (setq result2 (+ result2 1) )
        );if
        (setq count1 (+ count1 1) )
    );repeat
    result2
);defin
;;; ================================================
;                     DCL SECTION
;;; ================================================
(defun SETNEWNAME_DCL ( currentname1 errVal / tmpfpath1 infile1 dcl_id result1 oldBname1)
    (setq tmpfpath1 (strcat (getvar "MYDOCUMENTSPREFIX") "\\TEMP1DCL.DCL") )
    (setq infile1 (open tmpfpath1 "w") )
    (write-line 
        "dialog1 : dialog
            { label = \"Enter new name\";
                : text
                    {
                    key = \"old_name_field\";
                    } 
                : edit_box 
                    {
                     label = \"New name:\"; 
                     key = \"field1\"; 
                     mnemonic = \"W\"; 
                     width = 40; 
                     allow_accept = true; 
                    }
                : errtile
                    {
                     width = 50; 
                    }
                ok_cancel;  
            }"
        infile1
    );write line
    (close infile1)
    (setq dcl_id (load_dialog tmpfpath1) )
    (if (new_dialog "dialog1" dcl_id)
        (progn
            (set_tile "old_name_field" (strcat "Old name: " currentname1) )
            (if (= errVal "ON")
               (set_tile "error" "Error: Such name already exists in the drawing. Try different one")
            );if
            (set_tile "field1" currentname1 )
            (mode_tile "field1" 2);makes the field named "field1" active and lets you to start typing immediately
            (action_tile "accept" "(GetuInputNewName)")
          	(action_tile "cancel" "(DialogSetOldName) ")
            (start_dialog)
            (unload_dialog dcl_id)
            (vl-file-delete tmpfpath1 )
        );progn
        (exit);else
    );if
    result1
);defun

(defun GetuInputNewName()
    (setq result1 (get_tile "field1"))
    (done_dialog)
); defun

(defun DialogSetOldName()
    (setq result1 currentname1)
    (done_dialog)
); defun