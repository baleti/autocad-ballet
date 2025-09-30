;Quit all drawings without saving (CAD Studio - www.cadforum.cz)
;https://www.cadforum.cz/en/how-to-quit-all-open-drawings-without-saving-tip7723
(vl-load-com)
(defun C:close-all-without-saving ( / dwg)
 (vlax-for dwg (vla-get-Documents (vlax-get-acad-object))
  (if (= (vla-get-active dwg) :vlax-false)(vla-close dwg :vlax-false))
 )
 (command "._close" "_y")
)
