import pandas as pd
import os
import sys
from datetime import datetime


# Function to convert a date string to the OFX date format
def to_ofx_date(date):
    if isinstance(date, datetime):
        return date.strftime("%Y%m%d")
    else:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    return date_obj.strftime("%Y%m%d")


# Read transactions from xslx file

# REQUIRED pandas openpyxl

# Sample format of xslx file
# R1C1 - Bank ID
# R1C2 - Account ID
# R1C3 - Account Type (CREDITCARD, CHECKING, SAVINGS, MONEYMRKT)
# R1C4 - Currency (CAD, USD, MXN, EUR, GBR, CHF, JPY, INR, CNY, etc.)

# For CREDITCARD
#   Payments are positive numbers
#   Spending is negative numbers
#
# For CHECKING, SAVINGS, MONEYMRKT
#   Deposits are positive numbers
#   Spending is negative numbers


if len(sys.argv) != 2:
    print("USAGE: python csv-2-ofx.py <inputdata_name>.xlxs")
    sys.exit(1)


xlsx_file_path = sys.argv[1]  # Path to your XLSX file

output_file_path = os.path.splitext(xlsx_file_path)[0] + ".ofx"

df = pd.read_excel(xlsx_file_path, header=None)

# Extract account details from the first row
bank_id = df.iloc[0, 0]
acct_id = df.iloc[0, 1]
acct_type = df.iloc[0, 2]
currency_type = df.iloc[0, 3]

# Extract transaction data from the remaining rows
transactions = df.iloc[2:, [0, 1, 2]].values.tolist()
# Start creating the OFX content
ofx_content = f"""\
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
<SIGNONMSGSRSV1>
<SONRS>
<STATUS>
<CODE>0
<SEVERITY>INFO
</STATUS>
<DTSERVER>{datetime.now().strftime("%Y%m%d%H%M%S")}
<LANGUAGE>ENG
</SONRS>
</SIGNONMSGSRSV1>
<BANKMSGSRSV1>
<STMTTRNRS>
<TRNUID>1
<STATUS>
<CODE>0
<SEVERITY>INFO
</STATUS>
<STMTRS>
<CURDEF>{currency_type}
<BANKACCTFROM>
<BANKID>{bank_id}
<ACCTID>{acct_id}
<ACCTTYPE>{acct_type}
</BANKACCTFROM>
<BANKTRANLIST>
<DTSTART>{to_ofx_date(transactions[0][0])}
<DTEND>{to_ofx_date(transactions[-1][0])}
"""

# Adding each transaction
for date, payee, amount in transactions:
    amount = float(amount)
    if acct_type.upper() == "CREDITCARD":
        trntype = (
            "DEBIT" if amount > 0 else "CREDIT"
        )  # Charges are "DEBIT", Payments are "CREDIT"
    else:
        trntype = (
            "CREDIT" if amount > 0 else "DEBIT"
        )  # Deposits are "CREDIT", Withdrawals are "DEBIT"

    ofx_content += f"""\
<STMTTRN>
<TRNTYPE>{trntype}
<DTPOSTED>{to_ofx_date(date)}
<TRNAMT>{amount}
<FITID>{to_ofx_date(date)}{amount}
<NAME>{payee}
</STMTTRN>
"""

# Closing the OFX structure
ofx_content += f"""\
</BANKTRANLIST>
<LEDGERBAL>
<BALAMT>0.00
<DTASOF>{datetime.now().strftime("%Y%m%d%H%M%S")}
</LEDGERBAL>
</STMTRS>
</STMTTRNRS>
</BANKMSGSRSV1>
</OFX>
"""

# Save to an OFX file in the same directory with the same name as the input Excel file but with .ofx extension
with open(output_file_path, "w") as file:
    file.write(ofx_content)

print(f"OFX file created successfully: {output_file_path}")
