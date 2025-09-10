(defun c:save-viewports-to-files (/ doc desktop-path tab-name ss filename 
                                  old-ctab old-tilemode viewport-ent
                                  counter i all-objects temp-ss old-cmdecho
                                  j vp-on dxf-data cvport-id source-name base-name
                                  ps-width ps-height ms-center ms-height ms-width
                                  aspect-ratio boundary-pts obj-count ent-data
                                  min-x max-x min-y max-y test-pt found-inside
                                  found-outside sample-ents k)
  
  ;; Get source filename without path and extension
  (setq source-name (getvar "DWGNAME"))
  (setq base-name (vl-filename-base source-name))
  
  ;; Get desktop path
  (setq desktop-path (strcat (getenv "USERPROFILE") "\\Desktop\\"))
  
  ;; Store current settings
  (setq old-ctab (getvar "CTAB")
        old-tilemode (getvar "TILEMODE")
        old-cmdecho (getvar "CMDECHO"))
  
  ;; Turn off command echo
  (setvar "CMDECHO" 0)
  
  ;; Get all layout tabs
  (setq layouts (layoutlist))
  
  ;; Process each layout
  (foreach tab-name layouts
    (princ (strcat "\n\n################ PROCESSING LAYOUT: " tab-name " ################"))
    
    ;; Ensure we're in paper space for this layout
    (setvar "TILEMODE" 0)
    (setvar "CTAB" tab-name)
    
    ;; Get viewports in current layout
    (setq ss (ssget "X" (list '(0 . "VIEWPORT") 
                               (cons 410 tab-name)
                               '(-4 . "!=") '(69 . 1))))
    
    (if ss
      (progn
        (princ (strcat "\nFound " (itoa (sslength ss)) " viewport(s)"))
        
        ;; Create a selection set to collect all objects
        (setq all-objects (ssadd))
        
        ;; Process each viewport
        (setq i 0)
        (repeat (sslength ss)
          (setq viewport-ent (ssname ss i))
          (setq dxf-data (entget viewport-ent))
          (setq cvport-id (cdr (assoc 69 dxf-data)))
          
          (princ (strcat "\n\n=============== VIEWPORT " (itoa (1+ i)) " (ID: " (itoa cvport-id) ") ==============="))
          
          ;; Check if viewport is on
          (setq vp-on (if (assoc 90 dxf-data)
                        (= (logand (cdr (assoc 90 dxf-data)) 131072) 0)
                        t))
          
          (princ (strcat "\nStatus: " (if vp-on "ON" "OFF")))
          
          ;; Turn on if needed
          (if (not vp-on)
            (progn
              (princ "\nTurning viewport ON temporarily...")
              (setq dxf-data (subst (cons 90 (logand (cdr (assoc 90 dxf-data)) (~ 131072)))
                                   (assoc 90 dxf-data)
                                   dxf-data))
              (entmod dxf-data)
              (entupd viewport-ent)
              (setq dxf-data (entget viewport-ent))
            )
          )
          
          ;; Extract viewport parameters from DXF
          (setq ps-width (cdr (assoc 40 dxf-data)))
          (setq ps-height (cdr (assoc 41 dxf-data)))
          (setq ms-center (cdr (assoc 12 dxf-data)))
          (setq ms-height (cdr (assoc 45 dxf-data)))
          
          (princ "\n\n--- DXF DATA ANALYSIS ---")
          (princ (strcat "\nPaper Space Center (10): " (vl-princ-to-string (cdr (assoc 10 dxf-data)))))
          (princ (strcat "\nPaper Space Width (40): " (rtos ps-width 2 4)))
          (princ (strcat "\nPaper Space Height (41): " (rtos ps-height 2 4)))
          (princ (strcat "\nModel Space Center (12): " (vl-princ-to-string ms-center)))
          (princ (strcat "\nModel Space Height (45): " (rtos ms-height 2 4)))
          
          ;; Calculate model space width from aspect ratio
          (setq aspect-ratio (/ ps-width ps-height))
          (setq ms-width (* ms-height aspect-ratio))
          
          (princ (strcat "\nCalculated MS Width: " (rtos ms-width 2 4)))
          (princ (strcat "\nAspect Ratio: " (rtos aspect-ratio 2 4)))
          (princ (strcat "\nScale: 1:" (rtos (/ ms-height ps-height) 2 2)))
          
          ;; Calculate boundaries with 20% buffer
          (setq boundary-pts
            (list
              (list (- (car ms-center) (* ms-width 0.6))  ; 20% extra
                    (- (cadr ms-center) (* ms-height 0.6))
                    0.0)
              (list (+ (car ms-center) (* ms-width 0.6))
                    (+ (cadr ms-center) (* ms-height 0.6))
                    0.0)
            ))
          
          (princ "\n\n--- CALCULATED BOUNDARIES (with 20% buffer) ---")
          (princ (strcat "\nLower Left: " (vl-princ-to-string (car boundary-pts))))
          (princ (strcat "\nUpper Right: " (vl-princ-to-string (cadr boundary-pts))))
          (princ (strcat "\nWidth Range: " (rtos (car (car boundary-pts)) 2 2) " to " (rtos (car (cadr boundary-pts)) 2 2)))
          (princ (strcat "\nHeight Range: " (rtos (cadr (car boundary-pts)) 2 2) " to " (rtos (cadr (cadr boundary-pts)) 2 2)))
          
          ;; Switch to model space for selection
          (setvar "TILEMODE" 1)
          
          ;; DIAGNOSTIC: Sample some objects to see where they are
          (princ "\n\n--- DIAGNOSTIC: SAMPLING MODEL SPACE OBJECTS ---")
          (setq temp-ss (ssget "_X" '((410 . "Model"))))
          
          (if temp-ss
            (progn
              (princ (strcat "\nTotal objects in model space: " (itoa (sslength temp-ss))))
              
              ;; Sample first 10 objects to see their locations
              (princ "\nSample object locations (first 10):")
              (setq k 0)
              (setq found-inside 0)
              (setq found-outside 0)
              (setq min-x nil max-x nil min-y nil max-y nil)
              
              (repeat (min 10 (sslength temp-ss))
                (setq sample-ent (ssname temp-ss k))
                (setq ent-data (entget sample-ent))
                (setq test-pt (cdr (assoc 10 ent-data)))
                
                (if test-pt
                  (progn
                    ;; Track extents
                    (if (or (not min-x) (< (car test-pt) min-x)) (setq min-x (car test-pt)))
                    (if (or (not max-x) (> (car test-pt) max-x)) (setq max-x (car test-pt)))
                    (if (or (not min-y) (< (cadr test-pt) min-y)) (setq min-y (cadr test-pt)))
                    (if (or (not max-y) (> (cadr test-pt) max-y)) (setq max-y (cadr test-pt)))
                    
                    (princ (strcat "\n  " (cdr (assoc 0 ent-data)) " at " 
                                  (rtos (car test-pt) 2 2) "," 
                                  (rtos (cadr test-pt) 2 2)))
                    
                    ;; Check if inside bounds
                    (if (and (>= (car test-pt) (car (car boundary-pts)))
                             (<= (car test-pt) (car (cadr boundary-pts)))
                             (>= (cadr test-pt) (cadr (car boundary-pts)))
                             (<= (cadr test-pt) (cadr (cadr boundary-pts))))
                      (progn
                        (princ " [INSIDE]")
                        (setq found-inside (1+ found-inside))
                      )
                      (progn
                        (princ " [OUTSIDE]")
                        (setq found-outside (1+ found-outside))
                      )
                    )
                  )
                )
                (setq k (1+ k))
              )
              
              (princ (strcat "\nSample results: " (itoa found-inside) " inside, " (itoa found-outside) " outside"))
              
              (if (and min-x max-x min-y max-y)
                (progn
                  (princ "\n\n--- MODEL SPACE EXTENTS (from sample) ---")
                  (princ (strcat "\nX range: " (rtos min-x 2 2) " to " (rtos max-x 2 2)))
                  (princ (strcat "\nY range: " (rtos min-y 2 2) " to " (rtos max-y 2 2)))
                  
                  ;; Check if viewport bounds overlap with model extents
                  (princ "\n\n--- OVERLAP CHECK ---")
                  (if (or (> (car (car boundary-pts)) max-x)
                          (< (car (cadr boundary-pts)) min-x)
                          (> (cadr (car boundary-pts)) max-y)
                          (< (cadr (cadr boundary-pts)) min-y))
                    (princ "\n*** WARNING: Viewport bounds DO NOT overlap with model space objects! ***")
                    (princ "\nViewport bounds overlap with model space objects")
                  )
                )
              )
              
              ;; Now do actual selection with multiple methods
              (princ "\n\n--- SELECTION ATTEMPTS ---")
              
              ;; Method 1: Window selection
              (princ "\n1. Window selection...")
              (setq temp-ss (ssget "_W" (car boundary-pts) (cadr boundary-pts)))
              (if temp-ss
                (princ (strcat " Found " (itoa (sslength temp-ss)) " objects"))
                (princ " No objects")
              )
              
              ;; Method 2: Crossing selection
              (if (not temp-ss)
                (progn
                  (princ "\n2. Crossing selection...")
                  (setq temp-ss (ssget "_C" (car boundary-pts) (cadr boundary-pts)))
                  (if temp-ss
                    (princ (strcat " Found " (itoa (sslength temp-ss)) " objects"))
                    (princ " No objects")
                  )
                )
              )
              
              ;; Method 3: Manual filtering
              (if (not temp-ss)
                (progn
                  (princ "\n3. Manual filtering of all objects...")
                  (setq temp-ss (ssget "_X" '((410 . "Model"))))
                  (setq filtered-ss (ssadd))
                  (setq obj-count 0)
                  
                  (if temp-ss
                    (progn
                      (setq j 0)
                      (repeat (sslength temp-ss)
                        (setq test-ent (ssname temp-ss j))
                        (setq ent-data (entget test-ent))
                        (setq test-pt (cdr (assoc 10 ent-data)))
                        
                        (if test-pt
                          (if (and (>= (car test-pt) (car (car boundary-pts)))
                                   (<= (car test-pt) (car (cadr boundary-pts)))
                                   (>= (cadr test-pt) (cadr (car boundary-pts)))
                                   (<= (cadr test-pt) (cadr (cadr boundary-pts))))
                            (progn
                              (ssadd test-ent filtered-ss)
                              (setq obj-count (1+ obj-count))
                            )
                          )
                        )
                        (setq j (1+ j))
                      )
                      
                      (if (> obj-count 0)
                        (progn
                          (setq temp-ss filtered-ss)
                          (princ (strcat " Found " (itoa obj-count) " objects"))
                        )
                        (progn
                          (setq temp-ss nil)
                          (princ " No objects within bounds")
                        )
                      )
                    )
                  )
                )
              )
              
              ;; Add found objects to collection
              (if temp-ss
                (progn
                  (princ (strcat "\n\n*** SUCCESS: Adding " (itoa (sslength temp-ss)) " objects to export ***"))
                  (setq j 0)
                  (repeat (sslength temp-ss)
                    (ssadd (ssname temp-ss j) all-objects)
                    (setq j (1+ j))
                  )
                )
                (princ "\n\n*** ERROR: No objects found for this viewport ***")
              )
            )
            (princ "\nNo objects in model space!")
          )
          
          ;; Return to paper space
          (setvar "TILEMODE" 0)
          (setvar "CTAB" tab-name)
          
          ;; Restore viewport state if needed
          (if (not vp-on)
            (progn
              (setq dxf-data (entget viewport-ent))
              (setq dxf-data (subst (cons 90 (logior (cdr (assoc 90 dxf-data)) 131072))
                                   (assoc 90 dxf-data)
                                   dxf-data))
              (entmod dxf-data)
              (entupd viewport-ent)
            )
          )
          
          (setq i (1+ i))
        )
        
        ;; Export collected objects
        (if (> (sslength all-objects) 0)
          (progn
            (princ (strcat "\n\n################ EXPORTING " (itoa (sslength all-objects)) " OBJECTS ################"))
            
            ;; Switch to model space
            (setvar "TILEMODE" 1)
            
            ;; Create filename
            (setq filename (strcat desktop-path base-name "-" tab-name ".dwg"))
            
            ;; Make unique if exists
            (setq counter 1)
            (while (findfile filename)
              (setq filename (strcat desktop-path base-name "-" tab-name "_" (itoa counter) ".dwg"))
              (setq counter (1+ counter))
            )
            
            ;; Export using WBLOCK
            (command "_.WBLOCK" filename "" "0,0,0" all-objects "")
            
            ;; Return to paper space
            (setvar "TILEMODE" 0)
            
            (princ (strcat "\n\n*** FILE SAVED: " filename " ***"))
          )
          (princ "\n\n*** NO OBJECTS TO EXPORT FROM THIS LAYOUT ***")
        )
      )
      (princ (strcat "\nNo viewports found in layout: " tab-name))
    )
  )
  
  ;; Restore original settings
  (setvar "CTAB" old-ctab)
  (setvar "TILEMODE" old-tilemode)
  (setvar "CMDECHO" old-cmdecho)
  
  (princ "\n\n################################################")
  (princ "\nSaveViewportsToFiles completed!")
  (princ "\n################################################")
  (princ)
)

;; Visual test function
(defun c:draw-viewport-bounds (/ vp-ent dxf-data ms-center ms-height ms-width
                               ps-width ps-height aspect-ratio)
  
  (princ "\n\nDRAW VIEWPORT BOUNDS TEST")
  (princ "\nSelect a viewport to visualize its calculated bounds...")
  
  ;; Make sure we're in paper space
  (if (= (getvar "TILEMODE") 1)
    (setvar "TILEMODE" 0)
  )
  
  ;; Select viewport
  (setq vp-ent (car (entsel "\nSelect a viewport: ")))
  
  (if (and vp-ent (= (cdr (assoc 0 (entget vp-ent))) "VIEWPORT"))
    (progn
      (setq dxf-data (entget vp-ent))
      
      ;; Get parameters
      (setq ps-width (cdr (assoc 40 dxf-data)))
      (setq ps-height (cdr (assoc 41 dxf-data)))
      (setq ms-center (cdr (assoc 12 dxf-data)))
      (setq ms-height (cdr (assoc 45 dxf-data)))
      (setq aspect-ratio (/ ps-width ps-height))
      (setq ms-width (* ms-height aspect-ratio))
      
      (princ "\n\nViewport Data:")
      (princ (strcat "\n  MS Center: " (vl-princ-to-string ms-center)))
      (princ (strcat "\n  MS Width: " (rtos ms-width 2 2)))
      (princ (strcat "\n  MS Height: " (rtos ms-height 2 2)))
      
      ;; Switch to model space
      (setvar "TILEMODE" 1)
      
      ;; Draw boundary rectangle in RED
      (command "_.COLOR" "1")
      (command "_.RECTANGLE"
               (list (- (car ms-center) (/ ms-width 2))
                     (- (cadr ms-center) (/ ms-height 2)))
               (list (+ (car ms-center) (/ ms-width 2))
                     (+ (cadr ms-center) (/ ms-height 2))))
      
      ;; Draw center cross in YELLOW
      (command "_.COLOR" "2")
      (command "_.LINE"
               (list (- (car ms-center) (/ ms-width 10)) (cadr ms-center))
               (list (+ (car ms-center) (/ ms-width 10)) (cadr ms-center))
               "")
      (command "_.LINE"
               (list (car ms-center) (- (cadr ms-center) (/ ms-height 10)))
               (list (car ms-center) (+ (cadr ms-center) (/ ms-height 10)))
               "")
      
      ;; Draw buffered boundary (20% larger) in GREEN
      (command "_.COLOR" "3")
      (command "_.RECTANGLE"
               (list (- (car ms-center) (* ms-width 0.6))
                     (- (cadr ms-center) (* ms-height 0.6)))
               (list (+ (car ms-center) (* ms-width 0.6))
                     (+ (cadr ms-center) (* ms-height 0.6))))
      
      ;; Reset color
      (command "_.COLOR" "BYLAYER")
      
      (princ "\n\nDrawn in model space:")
      (princ "\n  RED rectangle = Calculated viewport bounds")
      (princ "\n  GREEN rectangle = Bounds with 20% buffer")
      (princ "\n  YELLOW cross = Center point")
      (princ "\n\nReturn to layout to check alignment...")
      
      ;; Return to paper space
      (setvar "TILEMODE" 0)
    )
    (princ "\nNo viewport selected")
  )
  (princ)
)

(princ "\n\n================================================")
(princ "\nVIEWPORT EXPORT WITH ENHANCED DIAGNOSTICS")
(princ "\n================================================")
(princ "\n")
(princ "\nCOMMANDS:")
(princ "\n  SAVEVIEWPORTSTOFILES - Export with detailed diagnostics")
(princ "\n  DRAWVIEWPORTBOUNDS - Draw viewport bounds in model space")
(princ "\n")
(princ "\nDIAGNOSTIC OUTPUT NOW INCLUDES:")
(princ "\n  - Sample of model space object locations")
(princ "\n  - Model space extents check")
(princ "\n  - Overlap verification")
(princ "\n  - Multiple selection method attempts")
(princ "\n")
(princ "\nThe diagnostics will show exactly why objects")
(princ "\nare or aren't being selected from each viewport.")
(princ)