import re
import pandas as pd
from dateutil.relativedelta import relativedelta
from helper_functions import IRR_calculation_clear_formula, ppmt, pmt, calculate_irr, write_IRR_results

# changeable local file path
file_path = './Loan IRR.xlsx'
sheet_IRR_calculation, sheet_charged_off, sheet_prepay = 'IRR Calculation', 'Charged Off', 'Prepay'


# clear formulas in loan data
IRR_calculation_clear_formula(file_path, sheet_IRR_calculation)


# read irr calculation data
df_loan_data = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_IRR_calculation)


# read and clean
df_charged_off = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_charged_off)
df_charged_off.set_index('Age', inplace=True)


# read and clean prepay data
df_prepay = pd.read_excel(file_path, engine='openpyxl', sheet_name=sheet_prepay, skiprows=1)
filtered_columns = df_prepay.columns[df_prepay.iloc[0].notna()]
df_prepay = df_prepay[filtered_columns]


# read static data
last_row = df_loan_data.last_valid_index()
static_variables = df_loan_data.iloc[:last_row + 1, 1:3].dropna()
variable_list = {}
for index, row in static_variables.iterrows():
    key = re.sub(r'[^a-zA-Z]', '', row[0])
    variable_list[key] = row[1]


# set static data to local variables
valuation_date = variable_list['ValuationDate']
grade = variable_list['Grade']
issue_date = variable_list['IssueDate']
term = variable_list['Term']
coupon_rate = variable_list['CouponRate']
invested = variable_list['Invested']
outstanding_balance = variable_list['OutstandingBalance']
recovery_rate = variable_list['RecoveryRate']
purchase_premium = variable_list['PurchasePremium']
servicing_fee = variable_list['ServicingFee']
earnout_fee = variable_list['EarnoutFee']
deafult_multiplier = variable_list['DeafultMultiplier']
prepay_multiplier = variable_list['PrepayMultiplier']

product_pos_name = str(term) + '-' + grade
product_pos = df_charged_off.columns.get_loc(product_pos_name)


# construct header of result
header = ['Months',
          'Paymnt_Count',
          'Paydate',
          'Scheduled_Principal',
          'Scheduled_Interest',
          'Scheduled_Balance',
          'Prepay_Speed',
          'Default_Rate',
          'Recovery',
          'Servicing_CF',
          'Earnout_CF',
          'Balance',
          'Principal',
          'Default',
          'Prepay',
          'Interest_Amount',
          'Total_CF']


# add index and header to result dataframe
num_rows = term + 1
data = {col: [None] * num_rows for col in header}
data['index'] = [i for i in range(term + 1)]
df = pd.DataFrame(data)
df = df.set_index('index')


# main calculation logic for each period
for i in df.index:
    # months
    df['Months'].at[i] = i + 1

    # payment count
    df['Paymnt_Count'].at[i] = i

    # payment date
    df['Paydate'].at[i] = issue_date + relativedelta(months=i)

    # scheduled principle
    df['Scheduled_Principal'].at[i] = ppmt(coupon_rate / 12, i, term, -invested) if i > 0 else 0

    # scheduled interest
    df['Scheduled_Interest'].at[i] = pmt(coupon_rate / 12, term, -invested) - df['Scheduled_Principal'].at[i] \
        if i > 0 else 0

    # scheduled balance
    df['Scheduled_Balance'].at[i] = df['Scheduled_Balance'].at[i - 1] - df['Scheduled_Principal'].at[i] \
        if i > 0 else invested

    # prepay speed
    df['Prepay_Speed'].at[i] = df_prepay[str(term) + 'M'][i - 1] if i > 0 else 0

    # default rate
    df['Default_Rate'].at[i] = df_charged_off[str(term) + '-' + grade][i + 1] if i + 1 < df_charged_off.shape[0] else 0

    # default
    df['Default'].at[i] = df['Balance'].at[i - 1] * df['Default_Rate'].at[i - 1] * deafult_multiplier \
        if i > 0 else 0

    # prepay
    temp_prepay = (df['Balance'].at[i - 1] - df['Scheduled_Interest'].at[i]) / df['Scheduled_Balance'].at[i - 1] \
                  * df['Scheduled_Principal'].at[i] \
        if i > 0 else 0
    df['Prepay'].at[i] = (df['Balance'].at[i - 1] - temp_prepay) * df['Prepay_Speed'].at[i] * prepay_multiplier \
        if i > 0 else 0

    # principal
    df['Principal'].at[i] = (df['Balance'].at[i - 1] - df['Default'].at[i]) / df['Scheduled_Balance'].at[i - 1] \
                            * df['Scheduled_Principal'].at[i] + df['Prepay'].at[i] \
        if i > 0 else 0

    # balance
    df['Balance'].at[i] = df['Balance'].at[i - 1] - df['Default'].at[i] - df['Principal'].at[i] if i > 0 else invested

    # recovery
    df['Recovery'].at[i] = df['Default'].at[i] * recovery_rate if i > 0 else 0

    # servicing cashflow
    df['Servicing_CF'].at[i] = (df['Balance'].at[i - 1] - df['Default'].at[i]) * servicing_fee / 12 if i > 0 else 0

    # earnout cashflow
    df['Earnout_CF'].at[i] = earnout_fee / 2 * invested if i == 12 or i == 18 else 0

    # interest amount
    df['Interest_Amount'].at[i] = (df['Balance'].at[i - 1] - df['Default'].at[i]) * coupon_rate / 12 if i > 0 else 0

    # total cashflow
    df['Total_CF'].at[i] = df['Principal'].at[i] + df['Interest_Amount'].at[i] + df['Recovery'].at[i] \
                           - df['Servicing_CF'].at[i] - df['Earnout_CF'].at[i] \
        if i > 0 else -invested * (1 + purchase_premium)


# reset index after calculation
df.reset_index(drop=True, inplace=True)


# calculate irr
irr_annualized = calculate_irr(df['Total_CF'].tolist()) * 12


# wirte result
write_IRR_results(file_path, sheet_IRR_calculation, df, irr_annualized)
