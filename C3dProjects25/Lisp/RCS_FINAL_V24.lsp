;;; ============================================================
;;; RCS_FINAL_V24.lsp
;;; Command: RCS_BUILD_CURVE_TABLE
;;;
;;; UPDATES:
;;; - Fixes "Too Few Arguments" error by making parsing safer.
;;; - Detects "C-1" style IDs (Relaxed ID matching).
;;; - Handles spaces after R= (e.g. "R= 200.00").
;;; ============================================================

(vl-load-com)

;;; --- CONFIGURATION ---
(setq RCS:ROW_HEIGHT 10.0)
(setq RCS:COL_WIDTHS '(12.0 38.0 20.0 20.0 20.0 20.0 25.0))
(setq RCS:TEXT_HEIGHT 2.5)
(setq RCS:MARGIN_H 0.50)
(setq RCS:MARGIN_V 0.25)

;;; --- MATH HELPERS ---
(defun RCS:Tan (ang) (if (not (equal (cos ang) 0.0 1e-6)) (/ (sin ang) (cos ang)) 0.0))
(defun RCS:DegToRad (d) (* d (/ pi 180.0)))

(defun RCS:ParseNum (s) 
  (if (and s (= (type s) 'STR))
    (distof (vl-string-trim " '" s))
    nil
  )
)

(defun RCS:DmsToDec (s / nums d m sec)
  (if (and s (= (type s) 'STR))
    (progn
      (setq s (vl-string-translate "°dD" "   " s))
      (setq s (vl-string-translate "'\"" "  " s))
      (setq nums (read (strcat "(" s ")")))
      (setq d (if (nth 0 nums) (float (nth 0 nums)) 0.0))
      (setq m (if (nth 1 nums) (float (nth 1 nums)) 0.0))
      (setq sec (if (nth 2 nums) (float (nth 2 nums)) 0.0))
      (+ d (/ m 60.0) (/ sec 3600.0))
    )
    nil
  )
)

(defun RCS:CalcTangent (R DeltaDec)
  (if (and (numberp R) (numberp DeltaDec)) 
      (* R (RCS:Tan (/ (RCS:DegToRad DeltaDec) 2.0))) 
      nil
  )
)

;;; --- STRING CLEANING ---

(defun RCS:Trim (s) (if (= (type s) 'STR) (vl-string-trim " \t\r\n" s) ""))

(defun RCS:StripMTextCodes (s / p1 p2)
  (setq s (vl-string-translate "\r\n" "  " s))
  (while (setq p1 (vl-string-search "{" s))
    (setq s (strcat (substr s 1 p1) (substr s (+ p1 2)))))
  (while (setq p1 (vl-string-search "}" s))
    (setq s (strcat (substr s 1 p1) (substr s (+ p1 2)))))
  (while (setq p1 (vl-string-search "\\" s))
     (setq p2 (vl-string-search ";" s p1))
     (if (and p2 (< (- p2 p1) 10)) 
       (setq s (strcat (substr s 1 p1) (substr s (+ p2 2))))
       (setq s (vl-string-subst " " "\\" s)) 
     )
  )
  (RCS:Trim s)
)

(defun RCS:StripTags (s)
  (if (= (type s) 'STR)
    (progn
      (setq s (vl-string-translate "()" "  " s))
      (foreach tag '("P&M" "P" "M" "A") (setq s (vl-string-subst "" tag s)))
      (RCS:Trim s)
    )
    ""
  )
)

;;; --- PARSING ENGINE ---

(defun RCS:MTextToLines (s / out start p seg)
  (setq s (vl-string-translate "\r" "\n" s))
  (while (setq p (vl-string-search "\\P" s)) (setq s (strcat (substr s 1 p) "\n" (substr s (+ p 3)))))
  (setq out '() start 1)
  (while (setq p (vl-string-search "\n" s (- start 1)))
    (setq seg (RCS:Trim (substr s start (- p (- start 1)))))
    (if (> (strlen seg) 0) (setq out (cons seg out)))
    (setq start (+ p 2))
  )
  (if (< start (+ (strlen s) 1)) (setq out (cons (RCS:Trim (substr s start)) out)))
  (reverse out)
)

(defun RCS:ParseCurveLabel (txt / lines id R_pm L_p L_m D_p D_m CB_p CD_p CB_m CD_m T_p T_m 
                                  cleanLine u val pPos mPos eqPos ePos wPos dirPos R_val D_p_val D_m_val nPos sPos startPos bear dist isPlat isMeas)
  (setq lines (RCS:MTextToLines txt))
  (setq id "" R_pm "" L_p "" L_m "" D_p "" D_m "" CB_p "" CD_p "" CB_m "" CD_m "" T_p "" T_m "")
  
  (foreach line lines
    (setq cleanLine (RCS:StripMTextCodes line))
    (setq u (strcase cleanLine))
    
    (cond
      ;; ID: Modified to catch C-1 (C*)
      ((wcmatch u "C*") (setq id cleanLine))
      
      ;; RADIUS
      ((wcmatch u "*R=*")
       (setq p (vl-string-search "R=" u)) 
       (if p (setq val (substr cleanLine (+ p 3))) (setq val cleanLine))
       (setq R_pm (RCS:StripTags val)))
       
      ;; LENGTH
      ((wcmatch u "*L=*")
       (setq p (vl-string-search "L=" u)) 
       (if p (setq val (substr cleanLine (+ p 3))) (setq val cleanLine))
       (setq pPos (vl-string-search "(P)" (strcase val))) 
       (setq mPos (vl-string-search "(M)" (strcase val)))
       (if (and pPos mPos)
         (progn (setq L_p (RCS:Trim (substr val 1 pPos))) (setq L_m (RCS:Trim (substr val (+ pPos 4) (- mPos (+ pPos 4))))))
         (setq L_p (RCS:StripTags val))))
      
      ;; DELTA
      ((and (or (wcmatch u "*=*") (wcmatch u "*\U+0394*") (wcmatch u "*D=*"))
            (not (wcmatch u "*R=*")) (not (wcmatch u "*L=*")))
       (setq eqPos (vl-string-search "=" cleanLine)) 
       (if (not eqPos) (setq eqPos (vl-string-search "\U+0394" cleanLine))) 
       (if (not eqPos) (setq eqPos 0))
       (setq val (substr cleanLine (+ eqPos 2)))
       (setq pPos (vl-string-search "(P)" (strcase val))) 
       (setq mPos (vl-string-search "(M)" (strcase val)))
       (if (and pPos mPos)
         (progn (setq D_p (RCS:Trim (substr val 1 pPos))) (setq D_m (RCS:Trim (substr val (+ pPos 4) (- mPos (+ pPos 4))))))))
      
      ;; CHORDS
      ((wcmatch u "*[NS]*[EW]*")
       (setq isPlat (wcmatch u "*(P)*")) 
       (setq isMeas (wcmatch u "*(M)*"))
       (setq ePos (vl-string-position (ascii "E") u)) 
       (setq wPos (vl-string-position (ascii "W") u)) 
       (setq dirPos (if ePos ePos wPos))
       (if dirPos
         (progn
           (setq nPos (vl-string-position (ascii "N") u)) 
           (setq sPos (vl-string-position (ascii "S") u))
           (setq startPos (if (and nPos sPos) (min nPos sPos) (if nPos nPos sPos))) 
           (if (not startPos) (setq startPos 0))
           (setq bear (substr cleanLine (+ startPos 1) (- (+ dirPos 1) startPos)))
           (setq dist (substr cleanLine (+ dirPos 2))) 
           (setq dist (RCS:StripTags dist))
           (if isPlat (setq CB_p bear CD_p dist)) 
           (if isMeas (setq CB_m bear CD_m dist)))))
    )
  )

  ;; MATH SAFETY
  (setq R_val (RCS:ParseNum R_pm))
  (setq D_p_val (RCS:DmsToDec D_p))
  (setq D_m_val (RCS:DmsToDec D_m))
  (if (and R_val D_p_val) (setq T_p (rtos (RCS:CalcTangent R_val D_p_val) 2 2)))
  (if (and R_val D_m_val) (setq T_m (rtos (RCS:CalcTangent R_val D_m_val) 2 2)))
  
  ;; ID FALLBACK
  (if (and (= id "") (> (length lines) 0))
      (if (wcmatch (RCS:StripMTextCodes (nth 0 lines)) "C*") (setq id (RCS:StripMTextCodes (nth 0 lines)))))

  (list 
    (cons 'ID id) 
    (cons 'R (strcat R_pm "\n(P&M)")) 
    (cons 'L (strcat L_p "(P)\n" L_m "(M)"))
    (cons 'DELTA (strcat D_p "(P)\n" D_m "(M)"))
    (cons 'CHB (strcat CB_p "(P)\n" CB_m "(M)"))
    (cons 'CHD (strcat CD_p "(P)\n" CD_m "(M)"))
    (cons 'T (strcat T_p "'(P)\n" T_m "'(M)"))
  )
)

;;; --- TABLE GENERATION ---

(defun RCS:BuildTable (recs inspt / tbl r rec rowCount colCount)
  (setq rowCount (+ 2 (length recs))) 
  (setq colCount 7)
  (setq tbl (vla-AddTable (vla-get-ModelSpace (vla-get-ActiveDocument (vlax-get-acad-object)))
                          (vlax-3D-point inspt) rowCount colCount RCS:ROW_HEIGHT 20.0))
  
  (vla-SetTextHeight tbl 1 RCS:TEXT_HEIGHT)
  (vla-SetTextHeight tbl 2 RCS:TEXT_HEIGHT)
  (vla-SetTextHeight tbl 4 RCS:TEXT_HEIGHT)
  (vla-put-HorzCellMargin tbl RCS:MARGIN_H)
  (vla-put-VertCellMargin tbl RCS:MARGIN_V)

  (vla-SetText tbl 0 0 "CURVE TABLE")
  (vla-MergeCells tbl 0 0 0 (- colCount 1))
  
  (vla-SetText tbl 1 0 "CURVE") (vla-SetText tbl 1 1 "CH BRG") (vla-SetText tbl 1 2 "CH DST")
  (vla-SetText tbl 1 3 "ARC") (vla-SetText tbl 1 4 "R") (vla-SetText tbl 1 5 "T") (vla-SetText tbl 1 6 "DELTA")
  (mapcar '(lambda (i w) (vla-SetColumnWidth tbl i w)) '(0 1 2 3 4 5 6) RCS:COL_WIDTHS)

  (setq r 2)
  (foreach rec recs
    (vla-SetText tbl r 0 (cdr (assoc 'ID rec)))
    (vla-SetText tbl r 1 (cdr (assoc 'CHB rec)))
    (vla-SetText tbl r 2 (cdr (assoc 'CHD rec)))
    (vla-SetText tbl r 3 (cdr (assoc 'L rec)))
    (vla-SetText tbl r 4 (cdr (assoc 'R rec)))
    (vla-SetText tbl r 5 (cdr (assoc 'T rec)))
    (vla-SetText tbl r 6 (cdr (assoc 'DELTA rec)))
    (foreach c '(0 1 2 3 4 5 6) (vla-SetCellAlignment tbl r c 5) (vla-SetCellTextHeight tbl r c RCS:TEXT_HEIGHT))
    (setq r (+ r 1))
  )
  (princ "\nTable generated successfully.")
)

(defun c:RCS_BUILD_CURVE_TABLE (/ ss i txt recs ins)
  (vl-load-com)
  (princ "\nSelect MText Curve Labels:")
  (setq ss (ssget '((0 . "MTEXT"))))
  (if ss
    (progn
      (setq recs '() i 0)
      (while (< i (sslength ss))
        (setq txt (vla-get-textstring (vlax-ename->vla-object (ssname ss i))))
        (setq recs (cons (RCS:ParseCurveLabel txt) recs))
        (setq i (+ i 1))
      )
      
      ;; Sort Numerically (Fixed for C-1 format)
      (setq recs (vl-sort recs (function (lambda (a b) 
         (< (atoi (vl-string-trim "C- " (cdr (assoc 'ID a)))) 
            (atoi (vl-string-trim "C- " (cdr (assoc 'ID b)))))))))
      
      (if (> (length recs) 0)
        (if (setq ins (getpoint "\nPick Table Insertion Point: ")) (RCS:BuildTable recs ins))
      )
    )
    (princ "\nNo MText selected.")
  )
  (princ)
)
