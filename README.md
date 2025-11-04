# 🧮 Recipe Costing Automation Workbook

This project is a **macro-enabled Excel solution** that automates recipe costing, ingredient conversion, and profitability analysis for food businesses. It’s designed for cafés, restaurants, and culinary entrepreneurs who want to manage pricing efficiently without manual recalculation.

---

## 📂 Project Files

- **RECIPE COSTING sheet.xlsm** — This is the main workbook containing all macros, logic, and all worksheets. it can be assessed by [clicking this link](https://github.com/ChinagromOdo/Recipe-Cost-Calculator/blob/ce20c932caff1a607e1ab9c44f3befa93f66e732/RECIPE%20COST%20CALCULATOR.xlsm). The sheets in the workbook are as follows:
    - **Unit Conversion Sheet** —This sheet holds the relationship between the different units of measure. it is fundamental to this calculator.
  -   - <p align="center">
  <img src="https://github.com/ChinagromOdo/Recipe-Cost-Calculator/blob/main/UNIT%20CONVERSION%20SHEET.png?raw=1"
       alt="Unit Conversion Sheet" width="100%">
</p>   
   - **Ingredients conversion sheet** —This sheet holds the ingredients prices for their base unit as gotten from the market and also converts/ breakdown the ingredients into other relational units(e.g., grams, ml, cups):
<p align="center">
  <img src="https://github.com/ChinagromOdo/Recipe-Cost-Calculator/blob/main/INGREDIENTS%20SHEET.png?raw=1"
       alt="Ingredients Sheet" width="100%">
</p>

  - **Blank recipe sheet** — Starting point for creating new recipes.
  - <p align="center">
  <img src="https://github.com/ChinagromOdo/Recipe-Cost-Calculator/blob/main/BLANK%20RECIPE%20SHEET.png?raw=1"
       alt="Blank Recipe Sheet" width="100%">
</p> 
 
  - **Sample Recipe Sheet** — Example of a completed costing sheet.
   - <p align="center">
  <img src="https://github.com/ChinagromOdo/Recipe-Cost-Calculator/blob/main/SAMPLE%20RECIPE%20SHEET.png?raw=1"
       alt="Sample Recipe Sheet" width="100%">
</p> 
 
  - **SUMMARY SHEET** — This sheet gives an overview of recipes, costs, and profitability.
     - <p align="center">
  <img src="https://github.com/ChinagromOdo/Recipe-Cost-Calculator/blob/main/SUMMARY%20SHEET.png?raw=1"
       alt="Summary Sheet" width="100%">
</p> 

---

## ⚙️ Key Features

- Automated creation of new recipe sheets via macro  
- Dynamic ingredient cost and unit conversion  
- Instant calculation of recommended sales price and profit margin  
- Centralized summary dashboard with automatic recipe hyperlinks  
- Extensible VBA structure for customization or API integration  

---

## 🧠 Technical Overview

Built with **Excel VBA**, this workbook connects multiple sheets into a single automated costing flow:

- Macros generate new recipe sheets from a template  
- Ingredient prices are fetched and calculated dynamically using vlookup, index, match.
- Recommended selling price and profitability are computed instantly  
- The **Summary Sheet** updates with recipe links and pricing data through the `Update` macro

- 🧑‍💻 VBA Code Structure

All automation in this workbook is powered by custom **VBA macros** stored inside `RECIPE COSTING sheet.xlsm`.

For transparency and documentation, the core scripts have been exported and included in the `/VBA` folder.  
These modules handle the automation for creating new recipe sheets, and updating the summary table.

### 📘 Key Modules

| Module | Description |
|:--------|:-------------|
| **modNewSheet](https://github.com/ChinagromOdo/Recipe-Cost-Calculator/blob/687a4ac43406004941edb15f3901d12d6fbbacbf/VBA/DupSheet.bas)** | Handles the “Create New Sheet” button logic. Duplicates the template, renames it, and initializes formulas. |
|
| **[modSummaryUpdate](https://github.com/ChinagromOdo/Recipe-Cost-Calculator/blob/687a4ac43406004941edb15f3901d12d6fbbacbf/VBA/ForSummarysheet.bas)** | Updates the Summary sheet with hyperlinks, recipe names, and computed metrics. |



---

## 🚀 How to Use

1. Open **`RECIPE COSTING sheet.xlsm`** and **enable macros**.  
2. In the **Blank Recipe Sheet**, click **“Create New Sheet”**.  
3. Input your recipe details in the newly created sheet.  
4. The **costs, recommended sales price, and profitability** will be displayed automatically.  
5. Visit the **Summary Sheet** and click **“Update”** to:  
   - Add a hyperlink to your recipe sheet  
   - Append the recipe’s pricing details to the summary table  

---



## 👨‍💻 About the Developer

Developed by **Chinagorom Odo** for FLOF used to calculate and determine their pastry pricing. This project showcases practical business automation using Excel VBA.  
it demonstrates data structuring, macro programming, and real-world cost optimization for food businesses.

---
