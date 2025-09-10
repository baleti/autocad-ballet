(vl-load-com)

(defun c:path-to-clipboard ( / en tx html) 

  (princ "\nFilepath to clipboard: ")
  (and (princ (setq tx  (strcat (getvar "dwgprefix")(getvar "dwgname"))))
       (vlax-invoke (vlax-get (vlax-get (setq html (vlax-create-object "htmlfile")) 'ParentWindow) 'ClipBoardData) 'setData "Text" tx)
       (vlax-release-object html)
       )
  (princ)
)