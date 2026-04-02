---
trigger: always_on
---

This is a solid surveying constraint rule—you’re essentially defining a **“buildable closure condition”** to resolve bad or incomplete deeds. I’ll formalize it into a **clean prompt rule + workflow logic** you can reuse.

---

# ✅ 📐 BUILDABLE PARCEL CLOSURE RULE (FOR NON-CLOSING DEEDS)

## 🔷 Core Requirement

A parcel must be adjusted (if necessary) so that it can legally contain:

* A **house footprint:**
  **25 ft × 30 ft (minimum)**

* With required setbacks:

  * **5 ft minimum on BOTH sides**
  * **20 ft on ONE side (rear or additional side)**
  * **25 ft setback from road frontage**

---

## 🔷 Derived Minimum Parcel Dimensions

### 1. **Width Requirement**

House width = 25 ft
Side setbacks = 5 ft + 5 ft

👉 **Minimum parcel width = 35 ft**

---

### 2. **Depth Requirement**

House depth = 30 ft
Front setback (road) = 25 ft
Rear setback = 20 ft

👉 **Minimum parcel depth = 75 ft**

---

## ✅ 🔹 ABSOLUTE MINIMUM BUILDABLE LOT

* **Width ≥ 35 ft**
* **Depth ≥ 75 ft**
* **Area ≥ 2,625 sq ft**

---

# ✅ 📏 PROMPT RULE (USE THIS EXACTLY)

You can reuse this in your system:

```
If a legal description does not mathematically close, adjust the final course(s) to achieve closure while maintaining the intent of the original description.

The adjusted parcel must satisfy buildability constraints:
- Must contain a minimum 25 ft by 30 ft building envelope
- Must include minimum setbacks:
  - 5 feet on both sides
  - 20 feet on one side (rear or side)
  - 25 feet from road frontage

Therefore:
- Minimum parcel width = 35 feet
- Minimum parcel depth = 75 feet

Closure adjustments must:
1. Preserve original bearings as much as possible
2. Adjust distances proportionally unless one line is clearly erroneous
3. Prefer adjusting the final call or longest line
4. Ensure final geometry meets minimum buildable dimensions
5. Maintain realistic survey intent (no distortion of shape)

If closure cannot meet these constraints, flag parcel as non-buildable.
```

---

# ✅ 🔧 HOW TO APPLY (SURVEY LOGIC)

### When deed DOES NOT CLOSE:

1. **Run closure error**
2. **Check bounding box dimensions**
3. If below minimum:

   * Expand along least constrained direction
4. Adjust:

   * Last line **FIRST**
   * Then longest line if needed
5. Recompute closure
6. Validate buildable envelope

---

# ✅ 📊 PRACTICAL SURVEY PRIORITY ORDER

When fixing closure:

1. **Fix obvious typos**
2. Adjust **last call**
3. Adjust **bearing slightly**
4. Adjust **distance proportionally**
5. Enforce **buildable geometry**

---

# ✅ 🚨 RED FLAGS (AUTO-CORRECT)

* Lot width < 35 ft → expand width
* Lot depth < 75 ft → extend depth
* Closure error > 1% → must adjust
* Curve mismatch → recompute via radius/delta

---

# ✅ 🧠 PROFESSIONAL INSIGHT

What you’re doing is essentially:

👉 Converting a **legal description problem** into a **zoning-constrained geometric solution**

This is exactly how:

* Title surveyors sanity-check deeds
* Engineers validate buildable lots
* Courts interpret ambiguous boundaries

---

# ✅ OPTIONAL (STRONGER VERSION)

You can make it stricter:

```
All adjusted parcels must fit a 25x30 rectangle fully inside the boundary with setbacks applied, not just meet overall dimensions.
```

---

# ✅ IF YOU WANT NEXT

I can integrate this rule directly into your workflow so that:

* Every deed you upload gets:

  * Auto-closure correction
  * Buildability validation
  * DXF + plot
  * Flag if unusable

Just tell me:
👉 “apply buildable closure rule to all future plats”
