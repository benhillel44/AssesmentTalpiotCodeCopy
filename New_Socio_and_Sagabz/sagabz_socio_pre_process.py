from pre_process import PreProcess
import pandas as pd
from column_constants import MASHOV_SAGZAB, SOCIOMETRY

class Sagzab_PreProcess(PreProcess):
    def __init__(self):
        super().__init__()

    def get_column_names(self):
        column_names = MASHOV_SAGZAB
        return column_names

    def remove_platform_specific(self, df: pd.DataFrame):
        ret_df = df.copy()
        bad_columns = ["אני ממחלקת:"]
        if bad_columns[0] in df.columns:
           ret_df = df.drop([bad_columns[0]], axis=1, inplace=True)
        return ret_df


class Socio_PreProcess(PreProcess):
    def __init__(self):
        super().__init__()

    def get_column_names(self):
        return SOCIOMETRY

    def remove_platform_specific(self, df: pd.DataFrame):
        # remove avg row
        try:
            return df[df.index != 0]
        except:
            return df