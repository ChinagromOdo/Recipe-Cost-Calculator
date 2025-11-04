# 🧮 Recipe Costing Automation Workbook

This project is a **macro-enabled Excel solution** that automates recipe costing, ingredient conversion, and profitability analysis for food businesses. It’s designed for cafés, restaurants, and culinary entrepreneurs who want to manage pricing efficiently without manual recalculation.

---

## 📂 Project Files

- **RECIPE COSTING sheet.xlsm** — Main workbook containing all macros and logic as well as all worksheets. the sheets in the workbook are as follows:
  - ***Ingredients conversion sheet** — Unit reference for ingredient standardization (e.g., grams, ml, cups):
  <p align="center">
  <img src="https://raw.githubusercontent.com/ChinagromOdo/Recipe-Cost-Calculator/main/Assets/INGREDIENTS%20SHEET.png" 
       alt="Ingredients Sheet" width="600">
</p>

  - **Blank recipe sheet** — Starting point for creating new recipes. [CLick to view sheet](https://github.com/ChinagromOdo/Recipe-Cost-Calculator/blob/e139511ea63110c4715b1c1f472c341bac8e1495/BLANK%20RECIPE%20SHEET.png)
  - **SAMPLE RECIPE SHEET** — Example of a completed costing sheet.  
  - **SUMMARY SHEET** — Overview of recipes, costs, and profitability.  
  - **UNIT CONVERSION SHEET** — Visual of the conversion system in use.  

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
