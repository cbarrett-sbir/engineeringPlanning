'''
This module provides a function to filter
out contract entries which have no hours
recorded to either of their 2 weeks
'''
import pandas as pd

def filterNaNs(df: pd.DataFrame) -> pd.DataFrame:
    '''
    Filter for identifying and dropping 
    contract rows where no hours have been
    entered for both weeks 1 & 2

    Params
    ------
        df: dataframe version of consolidated 
        excel time forecasts

    Returns
    -------
        a dataframe with contract rows without 
        hours dropped
    '''
    def _filter(df: pd.DataFrame) -> pd.DataFrame:
        '''Checks if all weekday columns are NaNs for both weeks'''  
        filtered_df = df

        weekdays = [
            'monday', 
            'tuesday', 
            'wednesday', 
            'thursday', 
            'friday'
        ]

        if all(df[weekdays].isna().all(axis=1)):
            # if all days are NaNs for weeks 1 & 2, drop those rows
            filtered_df = df[df['week'].eq(1) & df['week'].eq(2)]
        return filtered_df
    
    data = df.groupby(['contract', 'name']).apply(_filter).reset_index(drop=True)
    return data