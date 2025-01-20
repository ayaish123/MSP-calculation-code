# MSP-calculation-code
Code used to calculate the minimum sale price for r-PVC in (doi//).

### Description of the Code

This Python script performs a Monte Carlo simulation to calculate the Minimum Selling Price (MSP) of products derived from two types of waste materials: wire harness waste and tarpaulin waste. The script integrates probabilistic modeling, Excel automation, and data analysis to generate a large dataset of MSP values based on variable feedstock prices.

1. Feedstock Price Simulation:
   - The prices of wire harness and tarpaulin waste are generated using a truncated normal distribution, ensuring all values stay within specified bounds.
   - Key parameters like the mean, standard deviation, and bounds are defined to simulate 10,000 random price values for each waste stream.

2. Integration with Excel:
   - The script uses the xlwings library to interact with an Excel workbook containing economic and operational calculations.
   - Relevant feedstock prices are updated in specific cells within the OpEX Parameters worksheet.

3. Goal Seek in Excel:
   - A custom function automates Excels Goal Seek feature to calculate the MSP in the Economic Analysis worksheet. This ensures the MSP achieves a target value (e.g., zero profit/loss).
   - The computed MSP is stored along with the corresponding feedstock prices.

4. Results Storage:
   - The script collects MSP values and their corresponding feedstock prices into a Pandas DataFrame.
   - Results are exported to an Excel file or CSV for further analysis.

5. Performance Metrics:
   - The total runtime of the script is calculated and displayed.
   - Descriptive statistics for the generated MSP values are printed to summarize the results.

---

### Instructions for Using the Code

#### Prerequisites
- Python 3.8 or higher.
- Install the required Python libraries
  
- A properly configured Excel workbook with the following:
  - Sheets: 
    - OpEXParameters: Contains cells for feedstock prices (B3 for wire harness and B4 for tarpaulin waste).
    - Economic Analysis: Contains the target and changing cells for Goal Seek (L55 and U19 respectively).
  - Ensure the Excel workbook is saved in the same directory as the script with the correct filename (.xlsx).

---

#### Steps to Run the Script
1. Prepare the Excel Workbook:
   - Place the Excel file in the same directory as the script.
   - Confirm that the sheet names and cell references match the script.

2. Run the Script:
   - Execute the Python script:
     bash
     python script_name.py
     
   - The script will perform 10,000 simulations, updating Excel values and calculating MSP using Goal Seek.

3. View the Results:
   - Check the saved results file (e.g., results.xlsx) in the output directory.
   - The file contains a table with columns for:
     - Wire Harness Price
     - Tarpaulin Waste Price
     - MSP

4. Performance Metrics:
   - The script outputs the total runtime and descriptive statistics of the MSP values to the console.

---

### Notes
- The runtime of the script may vary depending on the system and Excels performance.
- Ensure that Excel’s macro security settings allow automation using xlwings.
- If needed, modify the cell references (L55, U19, etc.) in the goal_seek function to match your Excel workbooks layout.

---
