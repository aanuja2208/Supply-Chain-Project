# Manufacturing Inventory & Supplier Reliability Dashboard

## 📊 Overview
This project is an **Interview-Ready Inventory Planning Solution** designed for a Manufacturing Parts Network. It demonstrates core supply chain competencies: demand forecasting, statistical inventory optimization (ABC-XYZ), and supplier performance tracking.

---

## 🛠 Features
- **12-Month Simulated Dataset**: 60 SKUs across 3 plant locations with realistic lead times and demand variability.
- **ABC-XYZ Analysis**: Statistical classification based on both annual consumption value (ABC) and demand predictability (XYZ).
- **Safety Stock & ROP Optimization**: Automated calculations using statistical distribution formulas to prevent stockouts while minimizing working capital.
- **Supplier Scorecard**: On-time delivery (OTD%) and delay count tracking for 5 major suppliers.
- **What-If Scenario Analysis**: Dynamic modeling to predict the impact of demand surges (e.g., 20% increase) or lead-time delays on reorder points.

---

## 📂 Project Structure
- `Inventory_Planning_Project.xlsx`: The core application with 6 integrated sheets.
- `Manufacturing_Inventory_Raw_Data.csv`: The underlying synthetic dataset.
- `create_excel_project.py`: The Python logic used to generate the formulated workbook.

---

## 📈 Key Insights & Results
1. **Inventory Efficiency**: ABC-XYZ classification identifies 'AX' items for lean management and 'CZ' items for buffer stock.
2. **Operational Continuity**: Safety stock levels are optimized for a 95% service level.
3. **Supplier Risk**: Clear visibility into supplier delays, enabling data-driven procurement decisions.

---

## 🚀 How to Use
1. Download `Inventory_Planning_Project.xlsx`.
2. Open the **Dashboard** sheet for a high-level overview.
3. Use the **Scenario_Analysis** sheet to test impact of demand changes.
4. Refresh the **Raw_Data** if updating the underlying CSV.

---

## 🎓 Technical Demonstration
Built using:
- **Excel Power Query**: For data cleaning and transformation.
- **Advanced Formulas**: XLOOKUP, IFERROR, STDEV.P, SQRT, etc.
- **Pivot Tables & Charts**: For dynamic data visualization.
- **Python-to-Excel Logic**: Automated workbook generation.
