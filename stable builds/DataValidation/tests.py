'''
Contains tests to ensure the integrity 
of time forecast excel sheets
'''

import datetime
from typing import List

import pandas as pd
import openpyxl

def testContractValidity(series1: pd.DataFrame, contract_list: pd.DataFrame) -> List[str]:
    '''
    Get a list of strings from the first pandas Series
    that are not present in the second pandas Series.

    Parameters
    ----------
    - series1 (pd.Series): The first pandas Series.
    - filepath (str): path to contracts list

    Returns
    -------
        None
    '''

    s = list((set(series1) - set(contract_list)))
    return s

def weekContractsMatch(file_path: str, cn_week1: List) -> List[str]:
    cn_week2 = pd.read_excel(
        io=file_path,
        sheet_name="Plan",
        header=None,
        usecols="S",
        skiprows=range(0, 17),
        nrows=14,
        keep_default_na=True,
        names=["contract"]
    ).dropna()["contract"]

    return list((set(cn_week1) - set(cn_week2)))



def isValidName(name: str, valid_names: pd.Series) -> bool:
    ''''''
    return valid_names.str.contains(name).any()

def testSheetExistence(wb: openpyxl.Workbook) -> bool:
    ''''''
    if 'Plan' in wb.sheetnames:
        return True
    print('Sheet \"Plan\" does not exist')
    return False

def isCorrectDate(actual_date: datetime.datetime, date: datetime.datetime) -> bool:
    ''''''
    return date == actual_date
