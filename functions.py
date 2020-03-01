import pandas as pd
import numpy as np


def clean_df(df, org_column, extra_column):
    df.rename(columns={"In main currency": org_column}, inplace=True)
    df[extra_column] = np.float(0)
    df['Date'] = pd.to_datetime(df['Date'])
    df = df[['Date', 'Account', 'Category', 'Tags', 'ExpenseAmount', 'IncomeAmount', 'Description']]

    return df


def read_df(path, sheet_name):
    df = pd.read_excel(path,
                       sheet_name=sheet_name,
                       header=1,
                       usecols='A:D,G,I',
                       dtypeType={'B': object, 'C': object, 'D': object, 'I': object})

    return df
