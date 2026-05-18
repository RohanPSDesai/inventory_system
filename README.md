# MSME Inventory Consulting Diagnostic Suite Tool

## 📌 Project Overview
This tool generates an interactive, professional-grade diagnostic spreadsheet tailored for small-to-medium business owners (MSMEs). It bridges the gap between high-level financial accounting and everyday shop-floor operations by isolating complex formulas from the user data-entry plane.

The suite is split into two primary operational environments:
1. **Macro Health Calculators (Tab 1):** A data-driven model analyzing raw operating figures (COGS, physical stock values, lost orders) across currency perspectives, applying benchmark validation gates.
2. **Micro Concept Planner (Tab 2):** A qualitative self-assessment layout structured as standalone 'Business Scenario Cards' using plain English to deliver localized, step-by-step priority roadmaps.

---

## ⚙️ Core Architecture Mechanics

### 1. Dynamic Currency Selection Engine
* **Anchor Point:** Cell `C5` on Tab 1 serves as the global interface controller.
* **Logic Execution:** Cell `E5` parses this selection using a localized conditional lookup formula:
  `=IF(C5="INR", 83.50, 1.00)`
* **Downstream Integration:** All baseline inventory values typed in by the user are funneled through calculation cells that dynamically scale when the multiplier changes. This allows the business owner to view their metrics in either local rupees or international baseline dollars.

### 2. Isolated Mathematical Grading Scales
To evaluate operations objectively, raw performance fields are normalized into scored metrics using boundary limits based on industry standard performance:
* **Asset Turnover Score:** Evaluates your inventory cycles against standard targets:
  `=MIN(1, IF(Calculated_Turnover >= 4, 1, Calculated_Turnover / 4))`
* **Fulfillment Failure Score:** Grades operational continuity. It awards a perfect score if stockout errors stay below 2%, scaling down linearly as failure metrics approach 12%:
  `=MAX(0, IF(Stockout_Rate <= 0.02, 1, 1 - ((Stockout_Rate - 0.02) / 0.10)))`
* **System Shrinkage Score:** Measures inventory security, penalizing unknown losses that creep beyond a tight 1% margin:
  `=MAX(0, IF(Shrinkage_Rate <= 0.01, 1, 1 - ((Shrinkage_Rate - 0.01) / 0.05)))`

### 3. Automated Strategy Roadmaps
Tab 2 features a reactive action text matrix driven by logical branching. The spreadsheet reads the numeric indicator code from the user's selected business practice level and instantly overwrites the advisory block with tailored instructions:
`=IF(LEFT(D11,1)="0", "CRITICAL ACTION...", IF(LEFT(D11,1)="1", "IMPROVEMENT STEP...", "EXCELLENT..."))`

---

## 🛠️ Step-by-Step Deployment Instructions

### Step 1: Initialize Interactive Controls in Google Sheets
Once you open this generated spreadsheet file inside Google Sheets, you need to turn the static text anchors into interactive dropdown selectors:
1. Select cell **`C5`** on Tab 1. Go to **Insert > Dropdown** and add `USD` and `INR` as your primary options.
2. Navigate to Tab 2 and select cells **`D10`**, **`D15`**, **`D20`**, **`D25`**, **`D30`**, and **`D35`**.
3. Go to **Insert > Dropdown**, set your selection criteria to **Dropdown (from a list)**, and copy these three options exactly:
   * `0 - Checked by Look-and-Guess`
   * `1 - Tracked on Excel Spreadsheets`
   * `2 - System-Driven Routines`

### Step 2: Establish the Dynamic Currency Symbol Layer
To dynamically swap currency symbols (`$` vs `₹`) across the spreadsheet without modifying the underlying formulas, apply a conditional formatting rule:
1. Highlight your dynamic financial output columns on Tab 1 (**`D10:D13`** and **`D30:D34`**).
2. Click **Format > Conditional Formatting** from the main application menu.
3. Choose **Custom Formula is** from the condition rules menu and enter:
   ```excel
   ='Macro Health Calculators'!$C$5="INR"