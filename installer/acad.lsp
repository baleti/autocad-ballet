;; load lisp commands
(setq folder-path (strcat (getenv "APPDATA") "/autocad-ballet/"))
(foreach file (vl-directory-files folder-path "*.lsp")
  (load (strcat folder-path file)))

;; load invoker of dotnet commands
(command "netload" (strcat folder-path "InvokeAddinCommand.dll"))
