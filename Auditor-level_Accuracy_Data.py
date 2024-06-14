# Importing the necessary packages
import pandas as pd
import win32com.client as win32

# Creating outlook connection
olApp = win32.Dispatch("Outlook.Application")
# Reading base file.
input = r'C:\Users\skathir\Desktop\Test\feb.xlsx'
df1 = pd.read_excel(input)
# Taking the necessary columns.
df = pd.DataFrame(
    df1,
    columns=[
        'C_ASIN',
        'SA_Status',
        'C_Auditor login id',
        'SA_Risk Category',
        'SA_RCA(SIM) Id',
    ],
)
df['C_Auditor_login_id'] = df['C_Auditor login id'] + '@amazon.com'
df['SA_Risk_Category'] = df['SA_Risk Category']
df['asin'] = df['C_ASIN']
df['sim_id'] = df['SA_RCA(SIM) Id']

table = pd.DataFrame(data=df)

# Creating pivot
df_pivot = pd.pivot_table(
    table,
    index=['C_Auditor_login_id'],
    columns=['SA_Status'],
    values='asin',
    aggfunc='count',
    margins=True,
    fill_value=" ",
)
df_pivot.rename(columns={'All': 'Total Audits'}, inplace=True)

# Converting pivot to excel.
df_pivot.to_excel(r'C:\Users\skathir\PycharmProjects\super audit project\test1.xlsx')
new_table = pd.read_excel(
    r'C:\Users\skathir\PycharmProjects\super audit project\test1.xlsx', index_col=False
)
new_table.C_Auditor_login_id[new_table.C_Auditor_login_id == 'All'] = '@managers'


df_pivot1 = pd.pivot_table(
    table,
    index=[
        'C_Auditor_login_id',
    ],
    columns=['SA_Risk_Category'],
    values='asin',
    aggfunc='count',
    margins=True,
    fill_value=0,
)
df_pivot1.to_excel(r'C:\Users\skathir\PycharmProjects\super audit project\test2.xlsx')
new_table2 = pd.read_excel(
    r'C:\Users\skathir\PycharmProjects\super audit project\test2.xlsx', index_col=False
)

#Declare the accuracy metrics.
new_table2['High_N_Points'] = float(1.0)
new_table2['Low_N_Points'] = float(0.25)
new_table2['Medium_N_Points'] = float(0.5)
new_table2['Percentages'] = (
    1
    - (
        new_table2['Low'] * new_table2['Low_N_Points']
        + new_table2['Medium'] * new_table2['Medium_N_Points']
        + new_table2['High'] * new_table2['High_N_Points']
    )
    / new_table['Total Audits']
) * 100

Accuracy_table = pd.DataFrame(
    data=new_table2,
    columns=['C_Auditor_login_id', 'High', 'Medium', 'Low', 'Percentages'],
)
# print(Accuracy_table)

table_max = len(new_table)
table_row = 0
table_coloum = 1
#
# Read the relevant text file for the outlook body.
#
with open(
    r'C:\Users\skathir\PycharmProjects\super audit project\table style.txt'
) as file:
    table_style = file.read()

with open(
    r'C:\Users\skathir\PycharmProjects\super audit project\before table html.txt'
) as body_file:
    body_html = body_file.read()

with open(
    r'C:\Users\skathir\PycharmProjects\super audit project\after_table_html.txt'
) as body_file:
    last_body_html = body_file.read()

if table_max > table_row: # Check the condition for table len.
    for index, row in Accuracy_table.iterrows():
        id = row['C_Auditor_login_id'].replace('@amazon.com', '')
        error_per = row['Percentages']
        mail_id = row['C_Auditor_login_id']
        #
        #Convert DF to HTML format.
        table_html = new_table[table_row:table_coloum].to_html(table_id='DSM Report')
        table_html2 = Accuracy_table[table_row:table_coloum].to_html(
            table_id='Risk Report'
        )
        table_row += 1
        table_coloum += 1
        #outlook creation.
        mail_item = olApp.CreateItem(0)
        mail_item.To = 'skathir@amazon.com'
        mail_item.CC = 'skathir@amazon.com'
        mail_item.Subject = id + ' ' + 'Accuracy Dashboard'

        template_text = open('before table html.txt', 'r').read()
        final_mail_body = template_text.format(id, error_per)

        mail_item.HTMLBody = (
            final_mail_body
            + '<br/> SA Audit Table <br/>'
            + table_html
            + '<br/> Auditor Accuracy Table <br/>'
            + table_html2
            + last_body_html
        )
        mail_item.send
