;; Loc01 curb TLEN by corner — command-line report + typed prompt.
;; Load: (load "c:/repos/microstation-vba-project/scripts/TLEN01CORNER.lsp")
;; Run:  TLEN01CORNER

(defun tlen01-seg-len (ent / ed obj l)
  (setq ed (entget ent)
        obj (cdr (assoc 0 ed)))
  (cond
    ((= obj "LINE")
     (distance (cdr (assoc 10 ed)) (cdr (assoc 11 ed))))
    ((= obj "ARC")
     (* (cdr (assoc 40 ed))
        (if (minusp (setq l (- (cdr (assoc 51 ed)) (cdr (assoc 50 ed)))))
            (+ pi pi l) l)))
    (T 0.0))
)

(defun tlen01-mid (ent / ed p10 p11)
  (setq ed (entget ent)
        p10 (cdr (assoc 10 ed))
        p11 (cdr (assoc 11 ed)))
  (list (* 0.5 (+ (car p10) (car p11)))
        (* 0.5 (+ (cadr p10) (cadr p11))))
)

(defun tlen01-island (pt / x y)
  (setq x (car pt) y (cadr pt))
  (cond
    ((<= (distance pt '(1028470.0 187478.0)) 28.0) "NW")
    ((<= (distance pt '(1028564.0 187503.0)) 28.0) "N")
    ((<= (distance pt '(1028486.0 187421.0)) 28.0) "SW")
    ((<= (distance pt '(1028589.0 187445.0)) 28.0) "E")
    ((<= (distance pt '(1028501.0 187485.0)) 22.0) "MidN")
    ((<= (distance pt '(1028525.0 187428.0)) 22.0) "MidS")
    (T nil))
)

(defun tlen01-accum (lay key totals / ss i ent mid isl bucket L)
  (setq ss (ssget "X" (list (cons 0 "LINE") (cons 8 lay))))
  (if ss
    (progn
      (setq i 0)
      (repeat (sslength ss)
        (setq ent (ssname ss i)
              mid (tlen01-mid ent)
              isl (tlen01-island mid))
        (if isl
          (progn
            (setq L (tlen01-seg-len ent)
                  bucket (cdr (assoc isl totals))
                  ;; bucket = (S D C)
                  bucket
                  (cond
                    ((= key 0) (list (+ (nth 0 bucket) L) (nth 1 bucket) (nth 2 bucket)))
                    ((= key 1) (list (nth 0 bucket) (+ (nth 1 bucket) L) (nth 2 bucket)))
                    ((= key 2) (list (nth 0 bucket) (nth 1 bucket) (+ (nth 2 bucket) L)))
                  )
                  totals (subst (cons isl bucket) (assoc isl totals) totals)
            )
          )
        )
        (setq i (1+ i))
      )
    )
  )
  totals
)

(defun tlen01-print-one (name cid totals / b)
  (setq b (cdr (assoc name totals)))
  (princ (strcat
    "\n" name "  CornerID " cid
    "  Straight " (rtos (nth 0 b) 2 2)
    "  Depressed " (rtos (nth 1 b) 2 2)
    "  Corner " (rtos (nth 2 b) 2 2)
    "  TOTAL " (rtos (+ (nth 0 b) (nth 1 b) (nth 2 b)) 2 2)
  ))
)

(defun C:TLEN01CORNER (/ totals ans name)
  (setq totals
        (list
          (cons "NW"   '(0.0 0.0 0.0))
          (cons "N"    '(0.0 0.0 0.0))
          (cons "SW"   '(0.0 0.0 0.0))
          (cons "E"    '(0.0 0.0 0.0))
          (cons "MidN" '(0.0 0.0 0.0))
          (cons "MidS" '(0.0 0.0 0.0))
        ))
  (setq totals (tlen01-accum "Straight Steel Faced Concrete Curb (18'' Deep)" 0 totals))
  (setq totals (tlen01-accum "Depressed Steel Faced Concrete Curb (18'' Deep)" 1 totals))
  (setq totals (tlen01-accum "Corner Steel Faced Concrete Curb (18'' Deep)" 2 totals))

  (princ "\n==== Loc01 curb LF by corner ====")
  (tlen01-print-one "NW"   "1044031" totals)
  (tlen01-print-one "N"    "1044038" totals)
  (tlen01-print-one "SW"   "1044035" totals)
  (tlen01-print-one "E"    "1044037" totals)
  (tlen01-print-one "MidN" "1044033" totals)
  (tlen01-print-one "MidS" "1044035" totals)
  (princ "\n")

  ;; Always leave a type-in prompt (never modal alert-only)
  (initget 0 "NW N SW E MidN MidS")
  (setq ans (getkword "\nCorner detail [NW/N/SW/E/MidN/MidS] <Enter to finish>: "))
  (if ans
    (progn
      (setq name ans)
      (princ "\n---- detail ----")
      (cond
        ((= name "NW")   (tlen01-print-one "NW"   "1044031" totals))
        ((= name "N")    (tlen01-print-one "N"    "1044038" totals))
        ((= name "SW")   (tlen01-print-one "SW"   "1044035" totals))
        ((= name "E")    (tlen01-print-one "E"    "1044037" totals))
        ((= name "MidN") (tlen01-print-one "MidN" "1044033" totals))
        ((= name "MidS") (tlen01-print-one "MidS" "1044035" totals))
      )
      (princ "\n")
      (getstring "\nPress Enter (type anything if you want): ")
    )
  )
  (princ "\nTLEN01CORNER done.\n")
  (princ)
)

(princ "\nTLEN01CORNER loaded. Type TLEN01CORNER — results + type-in prompt on command line.\n")
(princ)
