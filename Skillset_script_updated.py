import pandas as pd
import numpy as np
from datetime import datetime
from uszipcode import SearchEngine
import pandas.io.formats.excel

pandas.io.formats.excel.header_style = None
pd.options.mode.chained_assignment = None  # default='warn'

branch_dict = {'Commerce, Inc.': 'B01', 'Commerce, LLC.': 'B02', 'Eagan': 'B03', 'Internal Payroll, Inc.': 'B04',
               'Internal Payroll, LLC.': 'B05', 'Paramount, Inc.': 'B06', 'Paramount, LLC.': 'B07',
               'Riverside, Inc.': 'B08', 'Riverside, LLC.': 'B09', 'Santa Ana, Inc.': 'B10',
               'Santa Ana, LLC.': 'B11', 'Santa Fe Springs, Inc.': 'B12',
               'Santa Fe Springs, LLC.': 'B13', 'SkillSet Group Inc - Oasis':
                   'B14', 'SkillSet Group LLC - Oasis': 'B15'}

missing_fields = ['MiddleName', 'Suffix', 'Address2', 'Occupation', 'DepartmentCode', 'Department', 'Division',
                  'LocationCode', 'Location', 'NewHire', 'NewHireEnrollByDate', 'BenefitsCalcDate', 'EffDateOverride']

states = ['AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DC', 'DE', 'FL', 'GA',
          'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD',
          'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
          'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC',
          'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY']


def main():
    dirty = input('File to Clean: ')

    clean_df = clean_file(dirty)

    format_sheet(clean_df)

    return


def format_sheet(df):
    """Formats the columns of the outputted spreadsheet
    to align with the standard census file."""

    # creating pandas excel writer using xlsxwriter and converting the data frame
    writer = pd.ExcelWriter("clean_census.xlsx", engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    format1 = workbook.add_format({'num_format': 'm/d/yyyy;@'})  # properly formatting our dates
    format2 = workbook.add_format({'num_format': '@'})  # formatting everything else to text

    # setting column formats
    worksheet.set_column('A:E', None, format2)
    worksheet.set_column('F:F', None, format1)
    worksheet.set_column('G:P', None, format2)
    worksheet.set_column('Q:Q', None, format1)
    worksheet.set_column('R:S', None, format2)
    worksheet.set_column('T:T', None, format1)
    worksheet.set_column('U:Z', None, format2)

    writer.save()


def find_zip(zip_code, state_bool=True, city_bool=True):
    """Returns state and city of zipcode the inputted. If state_bool
    is False state information will not be returned (likewise for
    city_bool). zip_code should be a string input."""

    search = SearchEngine(simple_zipcode=True)
    zipcode = search.by_zipcode(zip_code)

    if state_bool and city_bool:
        return zipcode.state, zipcode.city
    elif state_bool:
        return zipcode.state
    else:
        return zipcode.city


def clean_address(df):
    """Cleans the City and StateCode fields of the given data frame."""
    df_copy = df
    df['City'] = df_copy['City'].fillna('NaN')
    df['StateCode'] = df_copy['StateCode'].fillna('NaN')

    for record in range(len(df['City'])):
        if df['City'][record] == 'NaN':
            city = find_zip(df['ZipCode'][record], state_bool=False)
            df['City'][record] = city

    for record in range(len(df['StateCode'])):
        if (df['StateCode'][record] == 'NaN') or (df['StateCode'][record] not in states):
            state = find_zip(df['ZipCode'][record], city_bool=False)
            df['StateCode'][record] = state

    for status in range(len(df['City'])):
        if df['City'][status] == 'NaN':
            df['City'][status] = np.nan

    for status in range(len(df['StateCode'])):
        if df['StateCode'][status] == 'NaN':
            df['StateCode'][status] = np.nan


def clean_div(df):
    """Cleans division code, only input is the data frame."""
    df['DivisionCode'] = [None] * len(df)

    for record in range(len(df['EmployeeBranch'])):
        if 'Inc' in df['EmployeeBranch'][record]:
            df['DivisionCode'][record] = '0001'
        elif 'LLC' in df['EmployeeBranch'][record]:
            df['DivisionCode'][record] = '0002'
        else:
            df['DivisionCode'][record] = '0001'


def clean_rate(df):
    """Cleans the HourlyRate and AnnualPay columns,
    only input is the data frame."""
    df['HourlyRate'] = df['HourlyRate'].fillna('NaN')

    for status in range(len(df['HourlyRate'])):
        if (df['HourlyRate'][status] == 'NaN'):
            df['HourlyRate'][status] = 1
            df['AnnualPay'][status] = 2080.00

    for status in range(len(df['HourlyRate'])):
        if df['HourlyRate'][status] == 'NaN':
            df['HourlyRate'][status] = np.nan


def clean_laststatus(df):
    """Cleans LastStatusDate column based on the Status column.
    Input is the data frame."""
    df['LastStatusDate'] = df['LastStatusDate'].fillna('NaN')

    for status in range(len(df['Status'])):
        if (df['Status'][status] == 'A') and (df['LastStatusDate'][status] != 'NaN'):
            df['LastStatusDate'][status] = np.nan

    for status in range(len(df['Status'])):
        if (df['Status'][status] == 'T') and (df['LastStatusDate'][status] == 'NaN'):
            df['LastStatusDate'][status] = datetime.today().strftime('%m/%d/%Y')

    for status in range(len(df['Status'])):
        if df['LastStatusDate'][status] == 'NaN':
            df['LastStatusDate'][status] = np.nan


def format_times(df):
    """Properly formats columns with dates to ensure the final
    output of the excel file is in the correct format. Input is
    the data frame."""
    df['HireDate'] = pd.to_datetime(df['HireDate']).dt.date
    for val in range(len(df['HireDate'])):
        # if not isinstance(df['HireDate'][val], str):
        df['HireDate'][val] = df['HireDate'][val].strftime('%m/%d/%Y')

    for val in range(len(df['BirthDate'])):
        if not isinstance(df['BirthDate'][val], str):
            df['BirthDate'][val] = df['BirthDate'][val].strftime('%m/%d/%Y')

    for val in range(len(df['LastStatusDate'])):
        if not (isinstance(df['LastStatusDate'][val], str) or isinstance(df['LastStatusDate'][val], float)):
            df['LastStatusDate'][val] = df['LastStatusDate'][val].strftime('%m/%d/%Y')


def add_fields(df):
    """Adds all missing fields in order to output a properly
    formatted census file."""
    empt_list = [None] * len(df)

    for field in missing_fields:
        df[field] = empt_list


def clean_file(file):
    """Function bringing together all of the cleaning together.
    Input is the excel file to be cleaned and the output is the
    cleaned file in the form of a data frame."""

    df = pd.read_excel(file)  # read in dirty excel file

    # Clean SSN
    for record in range(len(df['SSN'])):
        df['SSN'][record] = df['SSN'][record].replace('-', '')

    # Clean Gender
    df['Gender'] = df['Gender'].fillna('F')

    # Clean MaritalStatusCode
    df['MaritalStatusCode'] = df['MaritalStatusCode'].fillna('S')

    # Clean BirthDate
    df['BirthDate'] = df['BirthDate'].fillna('01/01/1921')

    # Clean CountryCode
    Country = ['US'] * len(df)

    df['CountryCode'] = Country

    # Rename phone columns
    df = df.rename(columns={'EmployeeCellPhone': 'HomePhone'})

    clean_laststatus(df)

    # Clean LocationStateCode
    df['LocationStateCode'] = df['LocationStateCode'].fillna('CA')

    for record in range(len(df['PayCycle'])):
        df['PayCycle'][record] = 52

    # Clean Salaried
    df['Salaried'] = df['Salaried'].fillna('N')

    clean_rate(df)

    # Clean EmployeeBranchCode
    for branch in range(len(df['EmployeeBranch'])):
        df['EmployeeBranchCode'][branch] = branch_dict.get(df['EmployeeBranch'][branch])

    clean_div(df)

    clean_address(df)

    add_fields(df)

    format_times(df)

    # Dropping unused records (no SSN or on Internal Payroll, Inc.)
    for record in range(len(df['EmployeeBranch'])):
        if df['EmployeeBranch'][record] == 'Internal Payroll, Inc.':
            df.drop(labels=record, axis=0, inplace=True)

    df = df[df['SSN'].notna()]

    df.reset_index(drop=True, inplace=True)

    return df


if __name__ == '__main__':
    main()
