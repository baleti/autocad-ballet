(defun c:dps( / layout) ;toggle Display Plot Styles
(setq layout (vla-get-activelayout
(vla-get-activedocument (vlax-get-acad-object))))
(if (eq :vlax-true (vla-get-showplotstyles layout))
(vla-put-showplotstyles layout :vlax-false)
(vla-put-showplotstyles layout :vlax-true)
) ;if
(command "regenall")
;(vla-get-showplotstyles layout) ;test status
) ;defun