(defun c:color2layer (/ atts doc lay lays lokt)
(defun laycheck (ent color / lay)
(if (< 0 color 256)
(progn
(setq lay (vla-add lays (strcat "Color-" (itoa color))))
(vla-put-layer ent (vla-get-name lay))
(vla-put-color lay color)
(vla-put-color ent acbylayer)
)
)
)
(setq doc (vla-get-activedocument (vlax-get-acad-object))
lays (vla-get-layers doc)
)
(vlax-for lay lays
;;check for locked layers
(if (eq :vlax-true (vla-get-lock lay))
(progn
(setq lokt (cons (vla-get-name lay) lokt))
(vla-put-lock lay :vlax-false)
)
)
)
(vla-startundomark doc)
(vlax-for blk (vla-get-blocks doc)
(if (and (eq (vla-get-isxref blk) :vlax-false)
(not (vl-string-search "|" (vla-get-name blk)))
)
(progn
(vlax-for ent blk
(laycheck ent (vla-get-color ent))
(if (and (vlax-property-available-p ent "hasattributes")
(vla-get-hasattributes ent)
(setq atts (vlax-invoke ent "getattributes"))
)
(progn
(foreach att atts
(laycheck att (vla-get-color att))
(vla-update att)
)
)
)
)
)
)
)
(if lokt
;;reset locked layers
(foreach lay lokt
(vla-put-lock (vla-item lays lay) :vlax-true)
)
)
(vla-endundomark doc)
(princ "\nDone!")
(princ)
)