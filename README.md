# xlsx-2-ofx

## Converter for excel files to ofx format for import to accounting software

- created and tested using python 3.12.5

I did not want to upload my clients documents to a website or pay for a service when I had the transaction data but not official ofx files from the banks.

This was required to import into SAGE 50 DESKTOP in my case.

## Sample xls file included

### Sample format of xslx file

- R1C1 - Bank ID
- R1C2 - Account ID
- R1C3 - Account Type (CREDITCARD, CHECKING, SAVINGS, MONEYMRKT)
- R1C4 - Currency (CAD, USD, MXN, EUR, GBR, CHF, JPY, INR, CNY, etc.)

- For CREDITCARD
  - Payments are positive numbers
  - Spending is negative numbers

- For CHECKING, SAVINGS, MONEYMRKT
  - Deposits are positive numbers
  - Spending is negative numbers

## To Use

1. Create virtual environment
2. Enter virtual environment
3. Install requirements
    `pip install -r requirements.txt`
4. Run using
    `python3 xlsx-2-ofx.py <name_of_excelfile.xlsx>`
5. Output will be `<name_of_excelfile.ofx>` in the same directory as this is run.

## Info

- **ALWAYS BACK UP YOUR DATA BEFORE YOU IMPORT AN OFX FILE CREATED WITH THIS SCRIPT**
- will not be liable for any errors or problems incurred
- will only take data from the first tab, if you have a multi tab spreadsheet then move the tab you want to export to the first tab position
- the first 2 rows of the spreadsheet must be as in the sample
- Date must be in YYYY-MM-DD format (i.e. 2024-01-15)

## License

GPL-3.0 License
