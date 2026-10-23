;;; MidN flare ask — visible in AutoCAD (alert + getpoint + banner)
;;; Loaded by _acad_prompt_midn_flare_picks.py
(defun c:MIDNFLAREASK ( / tip left tan f banner path)
  (vl-load-com)
  (setq path "c:/repos/microstation-vba-project/scripts/_midn_flare_picks.txt")
  (alert
    (strcat
      "MidN NORTH yellow flare\n"
      "Click 3 corners of your blue sketch:\n"
      "1 TIP  = 35.01 (green topo meets tan)\n"
      "2 LEFT = green topo / DWS side\n"
      "3 TAN  = ON tan curb PAST green arrow"
    )
  )
  (command "_.ZOOM" "_C" "1028502,187483" "16")
  (setq banner
    (entmakex
      (list
        '(0 . "MTEXT")
        '(100 . "AcDbEntity")
        '(8 . "C-ANNO-NPLT")
        '(62 . 2)
        '(100 . "AcDbMText")
        '(10 1028490.0 187498.0 0.0)
        '(40 . 1.0)
        '(41 . 24.0)
        (cons 1
          (strcat
            "MidN NORTH flare — click TIP, then LEFT, then TAN\\P"
            "TIP=35.01  LEFT=green topo  TAN=tan past green arrow"
          )
        )
      )
    )
  )
  (setq tip (getpoint "\nAutoCAD 1/3 TIP at 35.01 (green meets tan): "))
  (if (null tip)
    (progn (if banner (entdel banner)) (princ "\nCANCELLED") (princ) (exit))
  )
  (alert "Next: 2/3 LEFT — green topo / DWS side of blue triangle")
  (setq left (getpoint tip "\nAutoCAD 2/3 LEFT along green topo / DWS: "))
  (if (null left)
    (progn (if banner (entdel banner)) (princ "\nCANCELLED") (princ) (exit))
  )
  (alert "Next: 3/3 TAN — ON tan curb PAST the green slope arrow")
  (setq tan (getpoint tip "\nAutoCAD 3/3 TAN on tan PAST green arrow: "))
  (if (null tan)
    (progn (if banner (entdel banner)) (princ "\nCANCELLED") (princ) (exit))
  )
  (if banner (entdel banner))
  (setq f (open path "w"))
  (write-line (strcat "tip," (rtos (car tip) 2 6) "," (rtos (cadr tip) 2 6)) f)
  (write-line (strcat "left," (rtos (car left) 2 6) "," (rtos (cadr left) 2 6)) f)
  (write-line (strcat "tan," (rtos (car tan) 2 6) "," (rtos (cadr tan) 2 6)) f)
  (close f)
  (princ "\nPICKS_OK — return to Cursor; solid will be placed.")
  (princ)
)
(c:MIDNFLAREASK)
