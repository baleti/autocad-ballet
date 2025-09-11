;; autocad-ballet startup loader
;; this file contains the main loading logic for AutoCAD Ballet

;; string replace function for trusted paths cleanup
(defun str-replace (old new str / pos)
  (setq pos 0)
  (while (setq pos (vl-string-search old str pos))
    (setq str (strcat (substr str 1 pos) new (substr str (+ pos (strlen old) 1))))
    (setq pos (+ pos (strlen new))))
  str)

;; generate random 28-character string for unique folder name
(defun generate-random-string (/ chars result seed i)
  (setq chars "abcdefghijklmnopqrstuvwxyz0123456789")
  (setq result "")
  (setq seed (fix (getvar "MILLISECS")))
  (setq i 0)
  (repeat 28
    (setq seed (rem (+ (* seed 7) i 17) 36))
    ;; ensure seed is always positive
    (if (< seed 0)
      (setq seed (+ seed 36)))
    (setq i (1+ i))
    (setq result (strcat result (substr chars (1+ seed) 1))))
  result)

;; check if this is AutoCAD LT
(defun is-autocad-lt (/ product)
  (setq product (getvar "PRODUCT"))
  (vl-string-search "LT" product))

;; parse trusted paths string into a list
(defun parse-trusted-paths (tp-string / result current-path semicolon-pos start)
  (setq result '())
  (setq start 1)
  (while (<= start (strlen tp-string))
    (setq semicolon-pos (vl-string-search ";" (substr tp-string start)))
    (setq current-path (if semicolon-pos
                         (substr tp-string start semicolon-pos)
                         (substr tp-string start)))
    (if (> (strlen current-path) 0)
      (setq result (cons current-path result)))
    (setq start (if semicolon-pos
                  (+ start semicolon-pos 1)
                  (+ (strlen tp-string) 1))))
  (reverse result))

;; join list of paths into semicolon-separated string
(defun join-trusted-paths (path-list / result)
  (setq result "")
  (foreach path path-list
    (setq result (if (= result "")
                   path
                   (strcat result ";" path))))
  result)

;; check if path is valid (absolute path with drive letter)
(defun is-valid-path (path)
  (or (vl-string-search ":\\" path)
      (vl-string-search ":/" path)))

;; main loader function
(defun ballet-load (/ appdata temp-base source-folder guid temp-folder fso tp
                     file-count dll-count dll-loaded folders-removed old-folder
                     old-folder-path delete-success lisp-folder lisp-count
                     path-list cleaned-paths)

  ;; set up paths
  (setq appdata (getenv "APPDATA"))

  ;; only load DLLs if not AutoCAD LT
  (if (not (is-autocad-lt))
    (progn
      ;; DLL loading logic for full AutoCAD
      (setq temp-base (strcat appdata "\\autocad-ballet\\commands\\bin\\temp\\"))

      ;; determine year from version
      (setq acad-version (substr (getvar "ACADVER") 1 4))
      (setq year-folder
        (cond
          ((= acad-version "21.0") "2017")
          ((= acad-version "22.0") "2018")
          ((= acad-version "23.0") "2019")
          ((= acad-version "23.1") "2020")
          ((= acad-version "24.0") "2021")
          ((= acad-version "24.1") "2022")
          ((= acad-version "24.2") "2023")
          ((= acad-version "24.3") "2024")
          ((= acad-version "25.0") "2025")
          ((= acad-version "25.1") "2026")
          (t "unknown")))

      (setq source-folder (strcat appdata "\\autocad-ballet\\commands\\bin\\" year-folder "\\"))

      ;; check if DLLs exist for this version
      (if (vl-file-directory-p source-folder)
        (progn
          (setq guid (generate-random-string))
          (setq temp-folder (strcat temp-base guid "\\"))

          (princ (strcat "\nAutoCAD Ballet: Session ID: " guid))

          ;; ensure parent directories exist
          (vl-mkdir (strcat appdata "\\autocad-ballet"))
          (vl-mkdir (strcat appdata "\\autocad-ballet\\commands"))
          (vl-mkdir (strcat appdata "\\autocad-ballet\\commands\\bin"))
          (vl-mkdir temp-base)

          ;; create new guid folder
          (vl-mkdir temp-folder)

          ;; create FSO object for file operations
          (setq fso (vlax-create-object "Scripting.FileSystemObject"))

          ;; copy files from source to temp
          (setq file-count 0)
          (foreach file (vl-directory-files source-folder "*.*" 1)
            (if (not (vl-catch-all-error-p
                       (vl-catch-all-apply 'vlax-invoke-method
                         (list fso 'CopyFile
                           (strcat source-folder file)
                           (strcat temp-folder file)
                           :vlax-true))))
              (setq file-count (1+ file-count))))
          (princ (strcat "\nAutoCAD Ballet: Copied " (itoa file-count) " files"))

          ;; add to trusted paths
          (setq tp (getvar "TRUSTEDPATHS"))
          (setq path-list (parse-trusted-paths tp))
          ;; add new temp folder if not already there
          (if (not (member temp-folder path-list))
            (setq path-list (cons temp-folder path-list)))
          (setvar "TRUSTEDPATHS" (join-trusted-paths path-list))

          ;; cleanup old temp folders AFTER setting up current one
          (if (vl-file-directory-p temp-base)
            (progn
              (setq tp (getvar "TRUSTEDPATHS"))
              (setq path-list (parse-trusted-paths tp))

              ;; first remove any corrupted partial paths (paths without drive letters)
              (setq cleaned-paths '())
              (foreach path path-list
                (if (is-valid-path path)
                  (setq cleaned-paths (cons path cleaned-paths))))
              (setq cleaned-paths (reverse cleaned-paths))

              ;; now remove old temp folders
              (setq folders-removed 0)
              (foreach dir (vl-directory-files temp-base nil -1)
                (if (and (not (member dir '("." "..")))
                         (/= dir guid)  ; never delete current folder
                         (vl-file-directory-p (strcat temp-base dir)))
                  (progn
                    (setq old-folder (strcat temp-base dir "\\"))
                    (setq old-folder-path (strcat temp-base dir))
                    ;; try to delete the folder using FSO
                    (setq delete-success nil)
                    (if (not (vl-catch-all-error-p
                               (vl-catch-all-apply 'vlax-invoke-method
                                 (list fso 'FolderExists old-folder-path))))
                      (if (not (vl-catch-all-error-p
                                 (vl-catch-all-apply 'vlax-invoke-method
                                   (list fso 'DeleteFolder old-folder-path :vlax-false))))
                        (setq delete-success t)))
                    ;; if deletion successful, remove from cleaned paths list
                    (if delete-success
                      (progn
                        (setq folders-removed (1+ folders-removed))
                        ;; remove this specific path from the list
                        (setq cleaned-paths
                              (vl-remove-if
                                '(lambda (p)
                                   (or (= (strcase p) (strcase old-folder))
                                       (= (strcase p) (strcase (substr old-folder 1 (1- (strlen old-folder)))))))
                                cleaned-paths)))))))

              ;; update trusted paths if any changes were made
              (if (or (> folders-removed 0)
                      (/= (length path-list) (length cleaned-paths)))
                (progn
                  (setvar "TRUSTEDPATHS" (join-trusted-paths cleaned-paths))
                  (if (> folders-removed 0)
                    (princ (strcat "\nAutoCAD Ballet: Cleaned " (itoa folders-removed) " old folders")))
                  (if (/= (length path-list) (length cleaned-paths))
                    (princ (strcat "\nAutoCAD Ballet: Removed "
                                  (itoa (- (length path-list) (length cleaned-paths)))
                                  " corrupted paths")))))))

          ;; release FSO
          (vlax-release-object fso)

          ;; load dll files
          (setq dll-count 0)
          (setq dll-loaded 0)
          (foreach dll (vl-directory-files temp-folder "*.dll" 1)
            (progn
              (setq dll-count (1+ dll-count))
              (setq dll-path (strcat temp-folder dll))
              (princ (strcat "\nAutoCAD Ballet: Loading " dll))
              (if (not (vl-catch-all-error-p
                         (vl-catch-all-apply 'vl-cmdf (list "netload" dll-path))))
                (setq dll-loaded (1+ dll-loaded)))))

          (if (> dll-count 0)
            (princ (strcat "\nAutoCAD Ballet: Loaded " (itoa dll-loaded) " of " (itoa dll-count) " DLL(s)"))
            (princ "\nAutoCAD Ballet: WARNING - No DLLs found in temp folder")))
        (princ "\nAutoCAD Ballet: No DLLs found for this AutoCAD version"))))

  ;; load lisp commands (for both full AutoCAD and LT)
  (setq lisp-folder (strcat appdata "\\autocad-ballet\\commands\\"))

  ;; Add lisp-folder to trusted paths before loading
  (setq tp (getvar "TRUSTEDPATHS"))
  (setq path-list (parse-trusted-paths tp))

  ;; remove corrupted paths again (in case we skipped DLL loading)
  (setq cleaned-paths '())
  (foreach path path-list
    (if (is-valid-path path)
      (setq cleaned-paths (cons path cleaned-paths))))
  (setq cleaned-paths (reverse cleaned-paths))

  ;; add commands folder if not already there
  (if (not (member lisp-folder cleaned-paths))
    (setq cleaned-paths (cons lisp-folder cleaned-paths)))

  (setvar "TRUSTEDPATHS" (join-trusted-paths cleaned-paths))

  (setq lisp-count 0)
  (foreach file (vl-directory-files lisp-folder "*.lsp")
    ;; skip loading startup-loader.lsp itself
    (if (not (= (strcase file) "STARTUP-LOADER.LSP"))
      (progn
        (load (strcat lisp-folder file))
        (setq lisp-count (1+ lisp-count)))))
  (if (> lisp-count 0)
    (princ (strcat "\nAutoCAD Ballet: Loaded " (itoa lisp-count) " LISP files")))

  (princ "\nAutoCAD Ballet: Ready\n")
  (princ))

;; run the loader
(ballet-load)
