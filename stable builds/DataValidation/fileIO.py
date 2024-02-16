import json

import pandas as pd

def getDefaultPaths(config_file_path: str):
    try:
        with open(config_file_path, 'r') as file:
            directory_paths = json.load(file)

            time_forecast_directory = directory_paths.get("time_forecast_directory", "")
            report_directory = directory_paths.get("report_directory", "")
            contracts_list = directory_paths.get("contracts_list_filepath", "")
            team_members_list = directory_paths.get("team_members_list_filepath", "")

            return time_forecast_directory, report_directory, contracts_list, team_members_list

    except FileNotFoundError:
        print(f"Config file '{config_file_path}' not found.")
        return None

def getContractList(PATH: str) -> pd.DataFrame:
    print("Reading ContractList.xlsx...")
    try:
        CONTRACT_LIST = pd.read_excel(PATH, "Sheet1", header=None)
    except Exception as e:
        print(e)
        return pd.DataFrame([])

    # create basis for contract info dataframe
    CONTRACT_LIST = CONTRACT_LIST[[0, 2, 1]][1:]
    CONTRACT_LIST = CONTRACT_LIST.set_axis(['contract', 'program_mgr', 'desc'], axis='columns')
    return CONTRACT_LIST

def getTeamList(PATH: str) -> pd.DataFrame:
    print("Reading TeamMembersList.xlsx...")
    try:
        team_list = pd.read_excel(PATH, "Sheet1", header=0)
    except Exception as e:
        print(e)
        return pd.DataFrame([])

    col_names = ['name', 'group', 'group_list', 'manager']
    team_list = team_list.set_axis(col_names, axis='columns')
    return team_list