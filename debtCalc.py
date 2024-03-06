### SIMPLE INTEREST CALCULATOR ###

## FORMULA: I = P * r * t
# I = interest, P = principal, r = annual interest rate, t = time the money is borrow in years

import pandas as pd
import math
import os
from datetime import datetime

def calc_Interest(P, r, t):
    I = P * r * t
    return I

print(calc_Interest(16196.67, .0689, 4))

### HOW LONG TO PAYOFF DEBT CALCULATOR (SIMPLE INTEREST) ###

## FORMULA: t = ln(PMT/PMT-r*P)/ln(1+r)
# t = time to pay off debt, PMT = payment amount, r = monthly interest rate(annual interest rate divided by 12), P is the principle

def calc_debt_payoff_simple(PMT, r, P):
    ''' returns payoff in months'''
    monthly_interest_rate = r / 12
    t = math.log(PMT / (PMT - monthly_interest_rate * P)) / math.log(1 + monthly_interest_rate)
    return t

print(calc_debt_payoff_simple(400, .0689, 5556.71))

### HOW LONG TO PAYOFF DEBT CALCULATOR (COMPOUND INTEREST) ###

## FORMULA: n = ln(PMT/PMT-r*P)/ln(1+r)
# n = number of payments, PMT = payment amount, r = monthly interest rate(per period), P is the principle

def calc_debt_payoff_compound(PMT, r, P):
    ''' returns payoff in months'''
    n = math.log(PMT / (PMT - (r / 12) * P)) / math.log(1 + r / 12)
    return n

print(calc_debt_payoff_compound(125, .2599, 5590.33))

### TURN MONTHS INTO YEARS FORMULA ###

def calc_month_to_year(f):
    ''' use simple or compound function for f'''
    y = f / 12
    return y

print(calc_month_to_year(calc_debt_payoff_compound(125, .2599, 5590.33)))

############################################## EXCEL SCRAPING ###########################################################################################

# file_path = 'E:\VSCode Files\Python Testing\Debt Table.csv'
# debt_table = pd.read_csv(file_path)

# print(debt_table)

# for index, row in debt_table.iterrows():
#     print(row)

# Load the Excel Sheet
df = pd.read_csv('E:\VSCode Files\Excel Exporting\Debt Table.csv')

#########################################################################################################################################################
############################################ CALCULATE TOTAL INTEREST PAID ##############################################################################
#########################################################################################################################################################

def calculate_compound_interest(principal, annual_rate):
    '''Return Compound Interest'''
    n = 12  # Compounded monthly
    t = 1   # Compounded for 1 year
    amount = principal * ((1 + (annual_rate / n)) ** (n * t))
    return amount

# Apply the compound interest calculation to each row
df['Compounded_Total'] = df.apply(lambda x: calculate_compound_interest(x['Total Owed'], x['Interest']), axis=1)

# Print the updated DataFrame
print(df[['Account', 'Compounded_Total']])

###########################################################################################################################################################
############################################# CALCULATE TOTAL MONTHS TO PAYOFF (COMPOUND INTEREST) ########################################################
###########################################################################################################################################################

def calculate_payoff_months(principal, annual_rate, payment):
    ''' Calculate Payoff in Months'''
    months = 0
    balance = principal
    monthly_interest_rate = annual_rate / 12
    
    # Check if the payment is sufficient to at least cover the first month's interest
    if payment < balance * monthly_interest_rate:
        return 'Infinity'
    
    while balance > 0:
        interest_for_month = balance * monthly_interest_rate
        balance = balance + interest_for_month - payment
        balance = max(0, balance)  # Ensure balance does not go negative
        months += 1
        if balance < payment:
            break
    
    return months

# Apply the updated function to the DataFrame
df['Payoff_Months'] = df.apply(
    lambda x: calculate_payoff_months(x['Total Owed'], x['Interest'], x['Min Payment']), axis=1
)

print(df[['Account', 'Payoff_Months']])

### EXPORT TO EXCEL ###

timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
output_file_path = f'E:\\VSCode Files\\Excel Exporting\\Debt Table Python Calc {timestamp}.xlsx'

df.to_excel(output_file_path, index=False)

print(f'The data with the payoff months has been saved to {output_file_path}')

# Open the Excel file
os.startfile(output_file_path)