# VBA-challenge

This VBA script is designed to analyze stock data for multiple years. It loops through all the stocks for one year and outputs the following information:

1. The ticker symbol.
2. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
3. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
4. The total stock volume of the stock.
5. Identifies and returns the stocks with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

**Files Included:**
   - `alphabetical_testing.xlsx`: This file contains less data and was used for testing purposes.
   - `Multiple_year_stock_data.xlsx`: This file contains the actual stock data for multiple years.

**Screenshots:**
   - Screenshots of each worksheet (.xlsx) are provided to confirm that the script works correctly. These screenshots serve as visual confirmation of the script's output.

## Notes:
- The `alphabetical_testing.xlsx` file was used for testing because it had less data than `Multiple_year_stock_data.xlsx`.
- The script is designed to work with any number of worksheets, each representing data for a different year.
- Adjustments have been made to the VBA script to ensure it runs on every worksheet in the workbook at once.