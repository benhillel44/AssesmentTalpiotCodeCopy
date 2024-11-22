from abc import ABC
import pandas as pd
import numpy as np

class Statistics(ABC):
    def __init__(self):
        pass

    def run(self, combined_data: pd.DataFrame):
        combined_data = combined_data.copy()
        # delete points to improve and points to conserve
        combined_data = combined_data.drop(columns=combined_data.columns[-2:])

        rows = []

        for df in combined_data.groupby("name"):
            person_name = df[0]
            df = df[1]
            columns = df.columns
            columns = columns.drop("name")
            for category in columns:
                col = df[category]
                # remove from col every non numeric value row
                col = col[col.apply(lambda x: isinstance(x, (int, np.int64, float, np.float64)))]
                mean = round(col.mean(), 2)
                std = round(col.std(), 2)
                row = [person_name, category, mean, std]
                rows.append(row)

        # what if I want to get more data, not open/close!!
        ret_val = pd.DataFrame(rows, columns=["name", "category", "mean", "std"])
        return ret_val