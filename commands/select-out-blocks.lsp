(defun C:select-out-blocks ()
  (if (ssget "_I"); there are preselected objects [= Implied selection]
    (sssetfirst nil (ssget "_I" '((0 . "INSERT")))); select/grip/highlight only Blocks among them
  )
)