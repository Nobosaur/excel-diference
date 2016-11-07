import pandas as pd

# Define the diff function to show the changes in each field
def report_diff(x):
    return x[0] if x[0] == x[1] else '{} ---> {}'.format(*x)

# We want to be able to easily tell which rows have changes
def has_change(row):
    if "--->" in row.to_string():
        return "Y"
    else:
        return "N"
def excel_differcence():
    # Read in both excel files
    df1 = pd.read_excel('C:\\Users\\KokotecN\\Documents\\python\\dev\\ExcelDiff\\sample_excel_1.xlsx', 'Sheet1', na_values=['NA'])
    df2 = pd.read_excel('C:\\Users\\KokotecN\\Documents\\python\\dev\\ExcelDiff\\sample_excel_2.xlsx', 'Sheet1', na_values=['NA'])

    # Make sure we order by account number so the comparisons work
    df1.sort_values(by="Stupac prvi")
    df1=df1.reindex()
    df2.sort_values(by="Stupac prvi")
    df2=df2.reindex()

    # Create a panel of the two dataframes
    diff_panel = pd.Panel(dict(df1=df1,df2=df2))

    #Apply the diff function
    diff_output = diff_panel.apply(report_diff, axis=0)

    # Flag all the changes
    diff_output['has_change'] = diff_output.apply(has_change, axis=1)

    #Save the changes to excel but only include the columns we care about
    diff_output[(diff_output.has_change == 'Y')].to_excel('my-diff-1.xlsx',index=False)

if __name__ == "__main__":
    excel_differcence()