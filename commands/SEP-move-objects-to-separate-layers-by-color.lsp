;SEP.lsp This program use to separate  layer with colour not bylayer
;Made by cad123, 1 Dec 02
(defun C:sep()
   
    (setq e(car(entsel "Pick On Grounp Layer To select:\n" )))
    (setq e(entget e))
    (setq n(cdr(assoc 8 e)))
    (setq nn(ssget "x" (list(cons 8 n))))
    (princ "\n Wait few minutes")
    (princ " ")
    (command "select" nn  "")
  
          
    (setq co 1)
    (while (< co 256)
      (setq ss (ssget "x" (list (cons 8 n)(cons 62 co))))
         (if (null ss) (setq co (1+ co))
    (progn
      (setq cname (itoa co))
      (setq lname(strcat  n cname))
      (command ".layer" "m" lname "c" co "" "")
      (command ".chprop" ss ""
           "C" "bylayer" "la" lname "")
        );end progn
    );end if
      );end while


  )