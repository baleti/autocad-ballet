;;  LayerQuellPick.lsp [command names: LOP, LOPN, LFP, LFPN]
;;  To "Quell" [turn Off or Freeze] Layers of Picked objects or Picked Nested objects.
;;  If selected object is on current Layer, changes current Layer to first one in Layer table
;;    that is neither off/frozen nor Xref-dependent nor that of selected object, and notifies
;;    User of new current Layer.  [Starting current Layer remains shown in Layer Properties
;;    Toolbar pull-down until command is terminated, but if Properties box is up, changes
;;    of current Layer are reflected immediately in its General section, Layer slot.]
;;  If selected nested object is part of a Block/Xref, and on Layer 0 within that reference's
;;    definition, turns off or freezes not Layer 0, but the deepest-nested non-0 Layer on which
;;    that reference or a containing reference is inserted, i.e. where nested object "appears."
;;  [Quirk I haven't found a way around yet: if Layer 0 is off or frozen, things on Defpoints
;;    Layer, even if visible, can't be selected; that Layer can't be quelled while Layer 0 is
;;    (using these commands, though it can be via Layer Manager or Toolbar).]
;;  Kent Cooper, November 2011

(defun C:LOP (); = Layer Off by Picking object(s) [top-level]
    (LQC entsel "" "turn Off" "_off")
); defun - C:LOP

(defun C:LOPN (); = Layer Off by Picking object(s) or Nested object(s)
    (LQC nentsel "/nested object" "turn Off" "_off")
); defun - C:LOPN

(defun C:LFP (); = Layer Freeze by Picking object(s) [top-level]
    (LQC entsel "" "Freeze" "_freeze")
); defun - C:LFP

(defun C:LFPN (); = Layer Freeze by picking object(s) or Nested object(s)
    (LQC nentsel "/nested object" "Freeze" "_freeze")
); defun - C:LFPN

(defun LQC (func pr1 pr2 opt / *error* cmde ent edata layq nestlist nlayq ldata layc)
  ; = Layer Quell with Current layer changed if necessary
  (defun *error* (errmsg); don't display error message if cancelled with Esc
    (if (not (wcmatch errmsg "Function cancelled,quit / exit abort,console break"))
      (princ (strcat "\nError: " errmsg))
    ); end if
    (setvar 'cmdecho cmde)
  ); defun - *error*
  (setq cmde (getvar 'cmdecho))
  (setvar 'cmdecho 0)
  (while
    (and
      (not ; to return T when (while) loop below is satisfied
        (while
          (and
            (not (setq ent (func (strcat "\nSelect object" pr1 " on Layer to " pr2 ": "))))
            (= (getvar 'errno) 7)
          ); and
          (prompt "\nNothing selected -- try again.")
        ); end while
      ); not
      ent ; something selected
    ); and
    (setq
      edata (entget (car ent))
      layq ; = LAYer to Quell [turn Off or Freeze]
        (if (= (cdr (assoc 8 edata)) "0"); on Layer 0
          (cond ; then
            ((> (length ent) 2); nested entity [other than Attribute] in Block/Xref
              (setq nestlist (last (nentselp (cadr ent)))); stack of references nested in
              (while
                (and
                  nestlist ; still nesting level(s) remaining
                  (= (setq nlayq (cdr (assoc 8 (entget (car nestlist))))) "0"); = Nested [non-0] LAYer to Quell
                ); and
                (setq nestlist (cdr nestlist)); move up a level if present
              ); while
              nlayq
                ; lowest-level non-0 Layer of nested or containing reference(s); 0 if that all the way up
            ); non-Attribute nested entity on Layer 0 condition
            ((= (cdr (assoc 0 edata)) "ATTRIB"); Attribute in Block
              (cdr (assoc 8 (entget (cdr (assoc 330 edata))))); Block's Layer
            ); Attribute on Layer 0 condition
            ("0"); none-of-the-above condition
          ); cond - then
          (cdr (assoc 8 edata)); else - Layer of entity/nested entity
        ); if & layq
    ); setq
    (if (= layq (getvar 'clayer)); selected object is on current Layer - find another
      (progn ; then
        (while ; look for Layer that's on, thawed, non-Xref-dependent, not selected object's
          (and
            (setq ldata (tblnext "layer" (not ldata))); still Layer(s) left in table
              ; nil if it gets to end of table without finding qualifying Layer
            (or
              (=
                (setq layc (cdr (assoc 2 ldata))); = LAYer to [possibly] make Current
                layq ; selected object's Layer
              ); =
              (minusp (cdr (assoc 62 ldata))); off [could be made current, but...]
              (= (logand (cdr (assoc 70 ldata)) 1) 1); frozen [can't be made current]
              (wcmatch layc "*|*"); Xref-dependent [can't be made current]
            ); or
          ); and
        ); while
        (if ldata ; found a qualifying Layer
          (setvar 'clayer layc); then - set it current
          (progn ; else - no qualifying Layer
            (alert "No available Layer to make current\nin place of selected object's Layer.")
            (quit)
          ); progn
        ); if
      ); progn
    ); if
    (command "_.layer" opt layq ""); turn off or freeze Layer of selected object
    (if ldata (prompt (strcat "\nCurrent Layer has been changed to " layc ".")))
  ); while
  (setvar 'cmdecho cmde)
  (princ)
); defun - LQC