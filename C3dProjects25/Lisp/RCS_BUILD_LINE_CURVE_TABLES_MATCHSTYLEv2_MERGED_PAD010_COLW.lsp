;;; ============================================================
;;; RCS_BUILD_LINE_CURVE_TABLES_MATCHSTYLEv2_MERGED_PAD010_COLW.lsp
;;; AutoCAD / Civil 3D
;;; Command: RCS_TABLES_FROM_WINDOW
;;;
;;; LINE TABLE:
;;; - 2 rows per line: (P or D) + A
;;; - Rule: you will either have (P)+(A/M) OR (D)+(A/M)
;;; - If any blank line fields, copy from the other source:
;;;     bearing-to-bearing, dist-to-dist (independently)
;;; - Column widths (fixed):
;;;     LINE=13.00, TYPE=13.00, BEARING=30.00, DIST=15.00
;;; - TYPE labels in cells: P / D / A
;;; - Text centered horiz + vert
;;; - Padding (ALL cells, ALL tables): Horz=0.10, Vert=0.10 (drawing units)
;;; - Non-annotative, plotted text height target: 0.10 (LINE TABLE)
;;;
;;; CURVE TABLE:
;;; - Unchanged parsing behavior for (P)/(M)/(P&M)
;;; - Uses dominant selected style + height
;;; - Applies same 0.10 table padding + centered cells
;;; ============================================================

(vl-load-com)

(if (not (boundp 'acAllViewports)) (setq acAllViewports 2))
(if (not (boundp 'acMiddleCenter)) (setq acMiddleCenter 5))

;;; ---------------------- GLOBAL CELL PADDING (ALL TABLES) ----------------------
(setq RCS:CELL_PAD 0.10)  ;; Horz + Vert cell margins (drawing units)

(defun RCS:SetTableMargins (tbl /)
  ;; Fixed padding for ALL tables/cells
  (if (vlax-property-available-p tbl 'HorzCellMargin) (vla-put-HorzCellMargin tbl RCS:CELL_PAD))
  (if (vlax-property-available-p tbl 'VertCellMargin) (vla-put-VertCellMargin tbl RCS:CELL_PAD))
)

;;; ---------------------- TARGET PLOTTED TEXT HEIGHT (NON-ANNOTATIVE) ----------------------
(setq RCS:PLOT_TEXT_H 0.10)   ;; plotted text height target (LINE TABLE)

(defun RCS:TargetModelHeight (plotH / cv ds)
  (setq cv (getvar "CVPORT"))
  (setq ds (getvar "DIMSCALE"))
  (if (or (null ds) (< ds 1.0)) (setq ds 1.0))
  (if (= cv 1) plotH (* plotH ds))
)

;;; ---------------------- STRING HELPERS ----------------------

(defun RCS:Trim (s)
  (if s (vl-string-trim " \t\r\n" s) "")
)

(defun RCS:IsBlank (s)
  (= (RCS:Trim (if s s "")) "")
)

(defun RCS:FirstNonBlank (a b)
  (if (not (RCS:IsBlank a)) a b)
)

(defun RCS:StrReplace (s find rep / p)
  (setq p (vl-string-search find s))
  (while p
    (setq s (strcat (substr s 1 p) rep (substr s (+ p (strlen find) 1))))
    (setq p (vl-string-search find s))
  )
  s
)

(defun RCS:RemoveMTextFormat (s / out i ch)
  ;; Minimal MTEXT formatting stripper
  (setq s (RCS:StrReplace s "\\P" "\n"))
  (setq s (RCS:StrReplace s "\\p" "\n"))
  (setq s (RCS:StrReplace s "{" ""))
  (setq s (RCS:StrReplace s "}" ""))

  (setq out "" i 1)
  (while (<= i (strlen s))
    (setq ch (substr s i 1))
    (if (= ch "\\")
      (progn
        (setq i (+ i 1))
        (if (<= i (strlen s))
          (progn
            (setq ch (substr s i 1))
            ;; Skip common MTEXT codes until ';'
            (if (member (strcase ch) '("H" "A" "C" "F" "Q" "T" "W"))
              (progn
                (while (and (<= i (strlen s)) (/= (substr s i 1) ";"))
                  (setq i (+ i 1))
                )
              )
            )
          )
        )
      )
      (setq out (strcat out ch))
    )
    (setq i (+ i 1))
  )
  out
)

(defun vl-string-position (char s start fromEnd / i)
  ;; Find index (0-based) of char in string; optional search from end
  (if (not start) (setq start 0))
  (if fromEnd
    (progn
      (setq i (strlen s))
      (while (and (> i 0) (/= (ascii (substr s i 1)) char))
        (setq i (- i 1))
      )
      (if (> i 0) (- i 1) nil)
    )
    (progn
      (setq i (+ start 1))
      (while (and (<= i (strlen s)) (/= (ascii (substr s i 1)) char))
        (setq i (+ i 1))
      )
      (if (<= i (strlen s)) (- i 1) nil)
    )
  )
)

;;; ---------------------- ENTITY HELPERS ----------------------

(defun RCS:GetEntText (e / obj oname)
  (setq obj (vlax-ename->vla-object e))
  (setq oname (strcase (vla-get-objectname obj)))
  (cond
    ((= oname "ACDBTEXT")
      (RCS:Trim (vla-get-textstring obj))
    )
    ((= oname "ACDBMTEXT")
      (RCS:Trim (RCS:RemoveMTextFormat (vla-get-textstring obj)))
    )
    (T "")
  )
)

(defun RCS:GetEntPoint (e / d)
  (setq d (entget e))
  (if (assoc 10 d) (cdr (assoc 10 d)) '(0.0 0.0 0.0))
)

(defun RCS:SortTextItems (items)
  ;; items: list of ( (x y z) "text" ) sorted top->bottom then left->right
  (vl-sort items
    (function
      (lambda (a b / pa pb xa ya xb yb)
        (setq pa (car a) pb (car b))
        (setq xa (car pa) ya (cadr pa))
        (setq xb (car pb) yb (cadr pb))
        (if (/= ya yb) (> ya yb) (< xa xb))
      )
    )
  )
)

;;; ---------------------- STYLE MATCHING ----------------------

(defun RCS:GetStyleFixedHeight (sty / td)
  (setq td (tblsearch "style" sty))
  (if td
    (if (assoc 40 td) (cdr (assoc 40 td)) 0.0)
    0.0
  )
)

(defun RCS:GetTextStyleAndHeight (e / obj oname sty h)
  (setq obj (vlax-ename->vla-object e))
  (setq oname (strcase (vla-get-objectname obj)))
  (setq sty "" h 0.0)

  (cond
    ((= oname "ACDBTEXT")
      (if (vlax-property-available-p obj 'StyleName) (setq sty (vla-get-StyleName obj)))
      (if (vlax-property-available-p obj 'Height)    (setq h (vla-get-Height obj)))
      (if (and (> (strlen sty) 0) (<= h 0.0)) (setq h (RCS:GetStyleFixedHeight sty)))
    )
    ((= oname "ACDBMTEXT")
      (cond
        ((vlax-property-available-p obj 'StyleName)     (setq sty (vla-get-StyleName obj)))
        ((vlax-property-available-p obj 'TextStyleName) (setq sty (vla-get-TextStyleName obj)))
      )
      (cond
        ((vlax-property-available-p obj 'TextHeight) (setq h (vla-get-TextHeight obj)))
        ((vlax-property-available-p obj 'Height)     (setq h (vla-get-Height obj)))
      )
      (if (and (> (strlen sty) 0) (<= h 0.0)) (setq h (RCS:GetStyleFixedHeight sty)))
    )
  )
  (list sty h)
)

(defun RCS:GetDominantStyleHeight (ss / i e sh sty h styles heights bestSty bestH)
  ;; Most common style; height = max height seen for that style.
  (setq styles '() heights '() i 0)
  (while (< i (sslength ss))
    (setq e (ssname ss i))
    (setq sh (RCS:GetTextStyleAndHeight e))
    (setq sty (car sh) h (cadr sh))

    (if (and sty (/= sty ""))
      (progn
        (if (assoc sty styles)
          (setq styles (subst (cons sty (+ 1 (cdr (assoc sty styles))))
                              (assoc sty styles)
                              styles))
          (setq styles (cons (cons sty 1) styles))
        )
        (if (> h 0.0)
          (if (assoc sty heights)
            (if (> h (cdr (assoc sty heights)))
              (setq heights (subst (cons sty h) (assoc sty heights) heights))
            )
            (setq heights (cons (cons sty h) heights))
          )
        )
      )
    )
    (setq i (+ i 1))
  )

  (setq bestSty "" bestH 0.0)
  (foreach p styles
    (if (or (= bestSty "") (> (cdr p) (cdr (assoc bestSty styles))))
      (setq bestSty (car p))
    )
  )
  (if (and bestSty (assoc bestSty heights)) (setq bestH (cdr (assoc bestSty heights))))

  (if (= bestSty "") (setq bestSty (getvar "TEXTSTYLE")))
  (if (<= bestH 0.0)
    (progn
      (setq bestH (RCS:GetStyleFixedHeight bestSty))
      (if (<= bestH 0.0) (setq bestH (getvar "TEXTSIZE")))
    )
  )
  (list bestSty bestH)
)

(defun RCS:SetCellTextStyleAny (tbl r c sty)
  (cond
    ((vlax-method-applicable-p tbl 'SetCellTextStyle) (vla-SetCellTextStyle tbl r c sty))
    ((vlax-method-applicable-p tbl 'SetTextStyle)     (vla-SetTextStyle tbl r c sty))
  )
)

(defun RCS:SetCellTextHeightAny (tbl r c h)
  (cond
    ((vlax-method-applicable-p tbl 'SetCellTextHeight) (vla-SetCellTextHeight tbl r c h))
    ((vlax-method-applicable-p tbl 'SetTextHeight)     (vla-SetTextHeight tbl r c h))
  )
)

(defun RCS:ApplyCellStyleHeight (tbl rows cols sty h / r c)
  (setq r 0)
  (while (< r rows)
    (setq c 0)
    (while (< c cols)
      (RCS:SetCellTextStyleAny tbl r c sty)
      (RCS:SetCellTextHeightAny tbl r c h)
      (setq c (+ c 1))
    )
    (setq r (+ r 1))
  )
)

(defun RCS:ApplyHeaderTitleSizes (tbl cols baseH / c)
  (setq c 0)
  (while (< c cols)
    (RCS:SetCellTextHeightAny tbl 0 c (* baseH 1.35))
    (setq c (+ c 1))
  )
  (setq c 0)
  (while (< c cols)
    (RCS:SetCellTextHeightAny tbl 1 c (* baseH 1.15))
    (setq c (+ c 1))
  )
)

;;; ---------------------- TABLE ALIGNMENT ----------------------

(defun RCS:CenterAllCells (tbl rows cols / r c)
  (if (vlax-method-applicable-p tbl 'SetCellAlignment)
    (progn
      (setq r 0)
      (while (< r rows)
        (setq c 0)
        (while (< c cols)
          (vla-SetCellAlignment tbl r c acMiddleCenter)
          (setq c (+ c 1))
        )
        (setq r (+ r 1))
      )
    )
  )
)

;;; ---------------------- LINE TABLE SPECIFIC COLUMN WIDTHS ----------------------

(defun RCS:SetLineTableColumnWidths (tbl /)
  ;; LINE=13, TYPE=13, BEARING=30, DIST=15
  (if (vlax-method-applicable-p tbl 'SetColumnWidth)
    (progn
      (vla-SetColumnWidth tbl 0 13.0)
      (vla-SetColumnWidth tbl 1 13.0)
      (vla-SetColumnWidth tbl 2 30.0)
      (vla-SetColumnWidth tbl 3 15.0)
    )
  )
)

;;; ---------------------- PARSING HELPERS ----------------------

(defun RCS:BeginsWith (s prefix)
  (= (strcase (substr s 1 (strlen prefix))) (strcase prefix))
)

(defun RCS:TagFromLine (s / u)
  ;; Returns: "P" "D" "M" "PM" "DM" or ""
  ;; - (A) treated as (M) for Actual
  ;; - (P&M)/(P&A) => "PM"
  ;; - (D&M)/(D&A) => "DM"
  (setq u (strcase s))
  (cond
    ((or (vl-string-search "(P&M" u)
         (vl-string-search "(P& M" u)
         (vl-string-search "(P &M" u)
         (vl-string-search "(P & M" u)
         (vl-string-search "(P&A" u)
         (vl-string-search "(P& A" u)
         (vl-string-search "(P &A" u)
         (vl-string-search "(P & A" u))
     "PM")
    ((or (vl-string-search "(D&M" u)
         (vl-string-search "(D& M" u)
         (vl-string-search "(D &M" u)
         (vl-string-search "(D & M" u)
         (vl-string-search "(D&A" u)
         (vl-string-search "(D& A" u)
         (vl-string-search "(D &A" u)
         (vl-string-search "(D & A" u))
     "DM")
    ((vl-string-search "(P)" u) "P")
    ((vl-string-search "(D)" u) "D")
    ((vl-string-search "(A)" u) "M")
    ((vl-string-search "(M)" u) "M")
    (T "")
  )
)

(defun RCS:StripTag (s / p)
  (setq p (vl-string-search "(" s))
  (if p (RCS:Trim (substr s 1 p)) (RCS:Trim s))
)

(defun RCS:LooksLikeBearingDist (s / u)
  (setq u (strcase (RCS:Trim s)))
  (or (wcmatch u "N * E *,N * W *,S * E *,S * W *")
      (wcmatch u "N*E*,N*W*,S*E*,S*W*"))
)

(defun RCS:ParseBearingDistance (s / clean tag lastsp)
  ;; Returns: (tag bearing dist)
  (setq tag (RCS:TagFromLine s))
  (setq clean (RCS:StripTag s))
  (setq lastsp (vl-string-position 32 clean nil T)) ;; last space
  (if lastsp
    (list tag
          (RCS:Trim (substr clean 1 lastsp))
          (RCS:Trim (substr clean (+ lastsp 2))))
    (list tag clean ""))
)

(defun RCS:ParseKeyValue (s / tag clean eqp key val)
  ;; Returns (tag key val)
  (setq tag (RCS:TagFromLine s))
  (setq clean (RCS:StripTag s))
  (setq eqp (vl-string-search "=" clean))
  (if eqp
    (progn
      (setq key (RCS:Trim (substr clean 1 eqp)))
      (setq val (RCS:Trim (substr clean (+ eqp 2))))
      (list tag key val))
    (list tag "" (RCS:Trim clean)))
)

(defun RCS:IsLineHeader (s / u)
  (setq u (strcase (RCS:Trim s)))
  (and (>= (strlen u) 2) (= (substr u 1 1) "L") (wcmatch u "L#*,L-#*"))
)

(defun RCS:IsCurveHeader (s / u)
  (setq u (strcase (RCS:Trim s)))
  (and (>= (strlen u) 2) (= (substr u 1 1) "C") (wcmatch u "C#*,C-#*"))
)

;;; ---------------------- RECORD HELPERS ----------------------

(defun RCS:NewLineRec (id)
  ;; P = Plat, D = Deed, M = Actual
  (list
    (cons 'ID id)
    (cons 'PBEAR "") (cons 'PDIST "")
    (cons 'DBEAR "") (cons 'DDIST "")
    (cons 'MBEAR "") (cons 'MDIST "")
  )
)

(defun RCS:NewCurveRec (id)
  (list
    (cons 'ID id)
    (cons 'PRAD "") (cons 'MRAD "")
    (cons 'PARC "") (cons 'MARC "")
    (cons 'PDELTA "") (cons 'MDELTA "")
    (cons 'PCHBRG "") (cons 'MCHBRG "")
    (cons 'PCHDST "") (cons 'MCHDST "")
  )
)

(defun RCS:SetRec (rec key val)
  (if (assoc key rec)
    (subst (cons key val) (assoc key rec) rec)
    (append rec (list (cons key val))))
)

(defun RCS:NormalizeLineRec (rec / pB pD dB dD aB aD refType refB refD)
  ;; Your rule: either (PLAT + ACTUAL) OR (DEED + ACTUAL)
  ;; Fill blanks from the other source (bearing->bearing, dist->dist)
  (setq pB (cdr (assoc 'PBEAR rec)))
  (setq pD (cdr (assoc 'PDIST rec)))
  (setq dB (cdr (assoc 'DBEAR rec)))
  (setq dD (cdr (assoc 'DDIST rec)))
  (setq aB (cdr (assoc 'MBEAR rec)))
  (setq aD (cdr (assoc 'MDIST rec)))

  ;; Decide ref: prefer PLAT if any plat value exists, else DEED if any deed value exists
  (if (or (not (RCS:IsBlank pB)) (not (RCS:IsBlank pD)))
    (setq refType "P")
    (if (or (not (RCS:IsBlank dB)) (not (RCS:IsBlank dD)))
      (setq refType "D")
      (setq refType "P")
    )
  )

  (cond
    ((= refType "P")
      (setq pB (RCS:FirstNonBlank pB aB))
      (setq pD (RCS:FirstNonBlank pD aD))
      (setq aB (RCS:FirstNonBlank aB pB))
      (setq aD (RCS:FirstNonBlank aD pD))
      (setq refB pB refD pD)
    )
    (T ;; "D"
      (setq dB (RCS:FirstNonBlank dB aB))
      (setq dD (RCS:FirstNonBlank dD aD))
      (setq aB (RCS:FirstNonBlank aB dB))
      (setq aD (RCS:FirstNonBlank aD dD))
      (setq refB dB refD dD)
    )
  )

  ;; Write back
  (setq rec (subst (cons 'PBEAR pB) (assoc 'PBEAR rec) rec))
  (setq rec (subst (cons 'PDIST pD) (assoc 'PDIST rec) rec))
  (setq rec (subst (cons 'DBEAR dB) (assoc 'DBEAR rec) rec))
  (setq rec (subst (cons 'DDIST dD) (assoc 'DDIST rec) rec))
  (setq rec (subst (cons 'MBEAR aB) (assoc 'MBEAR rec) rec))
  (setq rec (subst (cons 'MDIST aD) (assoc 'MDIST rec) rec))

  ;; Append computed ref info (REFTYPE is "P" or "D")
  (append rec (list (cons 'REFTYPE refType) (cons 'REFBEAR refB) (cons 'REFDIST refD)))
)

;;; ---------------------- TABLE HELPERS ----------------------

(defun RCS:TableAdd (ms inspt rows cols rowH colW / tbl)
  (vla-AddTable ms (vlax-3D-point inspt) rows cols rowH colW)
)

(defun RCS:SetCell (tbl r c txt)
  (vla-SetText tbl r c (if txt txt ""))
)

;;; ---------------------- LINE TABLE ----------------------

(defun RCS:MakeLineTable (lineRecs inspt sty selH / acad doc ms cols rows tbl rec h rowH colW rowIdx nrec id refType refB refD aB aD)
  (setq acad (vlax-get-acad-object))
  (setq doc  (vla-get-ActiveDocument acad))
  (setq ms   (vla-get-ModelSpace doc))

  ;; Layout: LINE | TYPE | BEARING | DIST
  (setq cols 4)
  ;; Two rows per line: REF then ACTUAL
  (setq rows (+ 2 (* 2 (length lineRecs))))

  ;; Forced plotted height (non-annotative)
  (setq h (RCS:TargetModelHeight RCS:PLOT_TEXT_H))

  ;; Keep your current height behavior; padding is table margin now (fixed 0.10)
  ;; Row height includes margin effect visually; keep simple.
  (setq rowH (+ h (* 2.0 RCS:CELL_PAD)))

  ;; Base column width for creation (overridden by per-column widths after)
  (setq colW 15.0)

  (setq tbl (RCS:TableAdd ms inspt rows cols rowH colW))

  ;; Title row
  (vla-SetText tbl 0 0 "LINE TABLE")
  (vla-MergeCells tbl 0 0 0 (- cols 1))

  ;; Header row
  (RCS:SetCell tbl 1 0 "LINE")
  (RCS:SetCell tbl 1 1 "TYPE")
  (RCS:SetCell tbl 1 2 "BEARING")
  (RCS:SetCell tbl 1 3 "DIST")

  ;; Data rows
  (setq rowIdx 2)
  (foreach rec lineRecs
    (setq nrec (RCS:NormalizeLineRec rec))

    (setq id      (cdr (assoc 'ID nrec)))
    (setq refType (cdr (assoc 'REFTYPE nrec))) ;; "P" or "D"
    (setq refB    (cdr (assoc 'REFBEAR nrec)))
    (setq refD    (cdr (assoc 'REFDIST nrec)))
    (setq aB      (cdr (assoc 'MBEAR nrec)))
    (setq aD      (cdr (assoc 'MDIST nrec)))

    ;; REF row (P or D)
    (RCS:SetCell tbl rowIdx 0 id)
    (RCS:SetCell tbl rowIdx 1 refType)
    (RCS:SetCell tbl rowIdx 2 refB)
    (RCS:SetCell tbl rowIdx 3 refD)
    (setq rowIdx (+ rowIdx 1))

    ;; ACTUAL row (A)
    (RCS:SetCell tbl rowIdx 0 id)
    (RCS:SetCell tbl rowIdx 1 "A")
    (RCS:SetCell tbl rowIdx 2 aB)
    (RCS:SetCell tbl rowIdx 3 aD)
    (setq rowIdx (+ rowIdx 1))
  )

  ;; Apply style + forced height
  (RCS:ApplyCellStyleHeight tbl rows cols sty h)
  (RCS:ApplyHeaderTitleSizes tbl cols h)

  ;; Padding + centered alignment
  (RCS:SetTableMargins tbl)
  (RCS:CenterAllCells tbl rows cols)

  ;; Apply LINE TABLE column widths
  (RCS:SetLineTableColumnWidths tbl)

  tbl
)

;;; ---------------------- CURVE TABLE (UNCHANGED PARSING; ADDS PADDING + CENTER) ----------------------

(defun RCS:MakeCurveTable (curveRecs inspt sty h / acad doc ms rows cols tbl r rec rowH colW)
  (setq acad (vlax-get-acad-object))
  (setq doc  (vla-get-ActiveDocument acad))
  (setq ms   (vla-get-ModelSpace doc))

  (setq cols 11)
  (setq rows (+ 2 (length curveRecs)))

  (setq rowH (* h 2.60))
  (setq colW (* h 18.0))

  (setq tbl (RCS:TableAdd ms inspt rows cols rowH colW))

  (vla-SetText tbl 0 0 "CURVE TABLE")
  (vla-MergeCells tbl 0 0 0 (- cols 1))

  (RCS:SetCell tbl 1 0  "CURVE")
  (RCS:SetCell tbl 1 1  "RADIUS (P)")
  (RCS:SetCell tbl 1 2  "RADIUS (M)")
  (RCS:SetCell tbl 1 3  "ARC (P)")
  (RCS:SetCell tbl 1 4  "ARC (M)")
  (RCS:SetCell tbl 1 5  "DELTA (P)")
  (RCS:SetCell tbl 1 6  "DELTA (M)")
  (RCS:SetCell tbl 1 7  "CH BRG (P)")
  (RCS:SetCell tbl 1 8  "CH DST (P)")
  (RCS:SetCell tbl 1 9  "CH BRG (M)")
  (RCS:SetCell tbl 1 10 "CH DST (M)")

  (setq r 2)
  (foreach rec curveRecs
    (RCS:SetCell tbl r 0  (cdr (assoc 'ID rec)))
    (RCS:SetCell tbl r 1  (cdr (assoc 'PRAD rec)))
    (RCS:SetCell tbl r 2  (cdr (assoc 'MRAD rec)))
    (RCS:SetCell tbl r 3  (cdr (assoc 'PARC rec)))
    (RCS:SetCell tbl r 4  (cdr (assoc 'MARC rec)))
    (RCS:SetCell tbl r 5  (cdr (assoc 'PDELTA rec)))
    (RCS:SetCell tbl r 6  (cdr (assoc 'MDELTA rec)))
    (RCS:SetCell tbl r 7  (cdr (assoc 'PCHBRG rec)))
    (RCS:SetCell tbl r 8  (cdr (assoc 'PCHDST rec)))
    (RCS:SetCell tbl r 9  (cdr (assoc 'MCHBRG rec)))
    (RCS:SetCell tbl r 10 (cdr (assoc 'MCHDST rec)))
    (setq r (+ r 1))
  )

  (RCS:ApplyCellStyleHeight tbl rows cols sty h)
  (RCS:ApplyHeaderTitleSizes tbl cols h)

  ;; Padding + center alignment (added)
  (RCS:SetTableMargins tbl)
  (RCS:CenterAllCells tbl rows cols)

  tbl
)

;;; ---------------------- MAIN PARSER ----------------------

(defun RCS:ParseSelectedTextToTables (items / lines curves mode curLine curCurve s tag bd kv val)
  (setq lines '() curves '() mode nil curLine nil curCurve nil)

  (foreach it items
    (setq s (cadr it))
    (if (= s "") (setq s nil))

    (if s
      (cond
        ((RCS:IsLineHeader s)
          (setq mode 'LINE)
          (setq curLine (RCS:NewLineRec (RCS:Trim s)))
          (setq lines (append lines (list curLine)))
        )

        ((RCS:IsCurveHeader s)
          (setq mode 'CURVE)
          (setq curCurve (RCS:NewCurveRec (RCS:Trim s)))
          (setq curves (append curves (list curCurve)))
        )

        (T
          (cond
            ;; ---------------- LINE CONTENT ----------------
            ((and (= mode 'LINE) curLine (RCS:LooksLikeBearingDist s))
              (setq bd (RCS:ParseBearingDistance s))
              (setq tag (car bd))
              (cond
                ;; PLAT + ACTUAL
                ((= tag "PM")
                  (setq curLine (RCS:SetRec curLine 'PBEAR (cadr bd)))
                  (setq curLine (RCS:SetRec curLine 'PDIST (caddr bd)))
                  (setq curLine (RCS:SetRec curLine 'MBEAR (cadr bd)))
                  (setq curLine (RCS:SetRec curLine 'MDIST (caddr bd)))
                )
                ;; DEED + ACTUAL
                ((= tag "DM")
                  (setq curLine (RCS:SetRec curLine 'DBEAR (cadr bd)))
                  (setq curLine (RCS:SetRec curLine 'DDIST (caddr bd)))
                  (setq curLine (RCS:SetRec curLine 'MBEAR (cadr bd)))
                  (setq curLine (RCS:SetRec curLine 'MDIST (caddr bd)))
                )
                ;; singles
                ((= tag "P")
                  (setq curLine (RCS:SetRec curLine 'PBEAR (cadr bd)))
                  (setq curLine (RCS:SetRec curLine 'PDIST (caddr bd)))
                )
                ((= tag "D")
                  (setq curLine (RCS:SetRec curLine 'DBEAR (cadr bd)))
                  (setq curLine (RCS:SetRec curLine 'DDIST (caddr bd)))
                )
                ((= tag "M") ;; Actual (A) or (M)
                  (setq curLine (RCS:SetRec curLine 'MBEAR (cadr bd)))
                  (setq curLine (RCS:SetRec curLine 'MDIST (caddr bd)))
                )
              )
              (setq lines (append (reverse (cdr (reverse lines))) (list curLine)))
            )

            ;; ---------------- CURVE CONTENT (UNCHANGED) ----------------
            ((and (= mode 'CURVE) curCurve)
              (setq tag (RCS:TagFromLine s))

              (cond
                ;; Radius
                ((or (RCS:BeginsWith (strcase (RCS:Trim s)) "R=")
                     (RCS:BeginsWith (strcase (RCS:Trim s)) "R ="))
                  (setq kv (RCS:ParseKeyValue s))
                  (setq val (caddr kv))
                  (cond
                    ((= tag "PM")
                      (setq curCurve (RCS:SetRec curCurve 'PRAD val))
                      (setq curCurve (RCS:SetRec curCurve 'MRAD val))
                    )
                    ((= tag "P") (setq curCurve (RCS:SetRec curCurve 'PRAD val)))
                    ((= tag "M") (setq curCurve (RCS:SetRec curCurve 'MRAD val)))
                  )
                )

                ;; Arc length (L=)
                ((or (RCS:BeginsWith (strcase (RCS:Trim s)) "L=")
                     (RCS:BeginsWith (strcase (RCS:Trim s)) "L ="))
                  (setq kv (RCS:ParseKeyValue s))
                  (setq val (caddr kv))
                  (cond
                    ((= tag "PM")
                      (setq curCurve (RCS:SetRec curCurve 'PARC val))
                      (setq curCurve (RCS:SetRec curCurve 'MARC val))
                    )
                    ((= tag "P") (setq curCurve (RCS:SetRec curCurve 'PARC val)))
                    ((= tag "M") (setq curCurve (RCS:SetRec curCurve 'MARC val)))
                  )
                )

                ;; Delta
                ((or (vl-string-search "Δ=" s)
                     (RCS:BeginsWith (strcase (RCS:Trim s)) "D=")
                     (RCS:BeginsWith (strcase (RCS:Trim s)) "DELTA="))
                  (setq kv (RCS:ParseKeyValue s))
                  (setq val (caddr kv))
                  (cond
                    ((= tag "PM")
                      (setq curCurve (RCS:SetRec curCurve 'PDELTA val))
                      (setq curCurve (RCS:SetRec curCurve 'MDELTA val))
                    )
                    ((= tag "P") (setq curCurve (RCS:SetRec curCurve 'PDELTA val)))
                    ((= tag "M") (setq curCurve (RCS:SetRec curCurve 'MDELTA val)))
                  )
                )

                ;; Chord bearing/dist
                ((RCS:LooksLikeBearingDist s)
                  (setq bd (RCS:ParseBearingDistance s))
                  (setq tag (car bd))
                  (cond
                    ((= tag "PM")
                      (setq curCurve (RCS:SetRec curCurve 'PCHBRG (cadr bd)))
                      (setq curCurve (RCS:SetRec curCurve 'PCHDST (caddr bd)))
                      (setq curCurve (RCS:SetRec curCurve 'MCHBRG (cadr bd)))
                      (setq curCurve (RCS:SetRec curCurve 'MCHDST (caddr bd)))
                    )
                    ((= tag "P")
                      (setq curCurve (RCS:SetRec curCurve 'PCHBRG (cadr bd)))
                      (setq curCurve (RCS:SetRec curCurve 'PCHDST (caddr bd)))
                    )
                    ((= tag "M")
                      (setq curCurve (RCS:SetRec curCurve 'MCHBRG (cadr bd)))
                      (setq curCurve (RCS:SetRec curCurve 'MCHDST (caddr bd)))
                    )
                  )
                )
              )

              (setq curves (append (reverse (cdr (reverse curves))) (list curCurve)))
            )
          )
        )
      )
    )
  )

  (list lines curves)
)

;;; ---------------------- COMMAND ----------------------

(defun c:RCS_TABLES_FROM_WINDOW (/ ss i e txt items sorted parsed
                                  lineRecs curveRecs insL insC
                                  acad doc sh selStyle selHeight)

  (setq acad (vlax-get-acad-object))
  (setq doc  (vla-get-ActiveDocument acad))

  (princ "\nSelect the TEXT/MTEXT block(s) with L# and C-# info (window/crossing is fine)...")
  (setq ss (ssget '((0 . "TEXT,MTEXT"))))

  (if (not ss)
    (progn
      (princ "\nNothing selected.")
      (princ)
    )
    (progn
      ;; Dominant style/height from selection (style used for both; height used for curve table)
      (setq sh (RCS:GetDominantStyleHeight ss))
      (setq selStyle (car sh))
      (setq selHeight (cadr sh))

      (princ (strcat
        "\nUsing Text Style: " selStyle
        " | Selected Height (curve): " (rtos selHeight 2 4)
        " | Line Target Height: " (rtos (RCS:TargetModelHeight RCS:PLOT_TEXT_H) 2 4)
        " | Cell Pad: " (rtos RCS:CELL_PAD 2 2)
        " | Line Cols: 13/13/30/15"
      ))

      ;; Collect items as ((x y z) "text")
      (setq items '() i 0)
      (while (< i (sslength ss))
        (setq e (ssname ss i))
        (setq txt (RCS:GetEntText e))
        (if (/= txt "")
          (setq items (cons (list (RCS:GetEntPoint e) txt) items))
        )
        (setq i (+ i 1))
      )

      (setq sorted (RCS:SortTextItems items))
      (setq parsed (RCS:ParseSelectedTextToTables sorted))
      (setq lineRecs (car parsed))
      (setq curveRecs (cadr parsed))

      (princ (strcat "\nParsed Lines: " (itoa (length lineRecs)) " | Curves: " (itoa (length curveRecs))))

      (setq insL (getpoint "\nPick insertion point for LINE TABLE: "))
      (if insL (RCS:MakeLineTable lineRecs insL selStyle selHeight))

      (setq insC (getpoint "\nPick insertion point for CURVE TABLE: "))
      (if insC (RCS:MakeCurveTable curveRecs insC selStyle selHeight))

      (vla-Regen doc acAllViewports)
      (princ "\nDone.")
      (princ)
    )
  )
)
