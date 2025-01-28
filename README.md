# VBA Stock Market Analysis  
**Christopher Kellam**  

**Languages and Tools Used in the Project:**  
- **Language:** VBA (Visual Basic for Applications)  
- **Tool:** Microsoft Excel  

## **Project Overview**  
This project utilized VBA scripting to analyze quarterly stock market data, automating calculations and generating insights into ticker performance. The script processes stock data to calculate quarterly changes, percentage changes, total stock volume, and highlights key metrics such as the greatest percentage increase, percentage decrease, and total volume.  

The features include:  
- Loops through stock data across all worksheets.  
- Outputs key metrics for each ticker symbol.  
- Applies conditional formatting to highlight positive and negative changes.  
- Dynamically processes multiple worksheets with a single script.  

---

## **1. Data Source**  
The project uses sample stock market data stored in Excel worksheets. Each worksheet represents a single quarter of data, enabling analysis across time periods.  

---

## **2. Functionality**  

### **Key Metrics Calculated:**  
1. **Ticker Symbol:** The unique identifier for each stock.  
2. **Quarterly Change:** The difference between the opening price at the start of a quarter and the closing price at the end.  
3. **Percentage Change:** The percentage difference between the opening and closing prices.  
4. **Total Stock Volume:** The total volume of stocks traded during the quarter.  

### **Enhanced Insights:**  
- Identifies the stocks with:  
  - The greatest percentage increase.  
  - The greatest percentage decrease.  
  - The greatest total volume.  

### **Automation Features:**  
- The script runs across all worksheets in the workbook, processing each quarter’s data sequentially.  
- Conditional formatting highlights:  
  - Positive changes in green.  
  - Negative changes in red.  

---

## **3. Example Output**   
Displays quarterly stock data with calculated columns for:  
- Ticker Symbol  
- Quarterly Change ($)  
- Percent Change (%)  
- Total Stock Volume  

![Example](Quarter 1.PNG)  

---

## **4. Submission Requirements**  
The project includes:  
- Screenshots of results demonstrating script outputs.  
- VBA script files for the analysis.  
- This README file detailing the project.  

---