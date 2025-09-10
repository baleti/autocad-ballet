;(C) 2010 CAD Studio - www.cadstudio.cz
;modified 24.11.22 Daniel Krajnik
;don't know how to use existing selection or select multiple by crossing window rather than picking them one by one this is stupid who wrote this arghh
;select all blocks like selected block(s)

(defun C:select-same-block ( / e blk ss blkl)
 (setq blkl ""  ss (ssadd))
 (while (setq e (entsel "\nSelect next block name to select <exit>: "))
  (setq blkl (strcat blkl "," (cdr (assoc 2 (entget (car e))))))
  (princ (substr blkl 2))
 );while
 (setq blkl (substr blkl 2))
 (if (> blkl "") (setq ss (ssget "_X" (list (cons 2 blkl)(cons 0 "INSERT")))))
 (if (zerop (getvar "CMDACTIVE"))
  (progn (sssetfirst ss ss)(princ "Use 'P' for this selection set: ")(princ))
   ss
 )
)
