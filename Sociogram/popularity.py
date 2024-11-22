import numpy as np
import pandas as pd

# Thresholds
LOW_THRESH = 3
HIGH_THRESH = 10
DIFF_THRESH = 5

OUTPUT_EXCEL = 'output_table.xlsx'
OUTPUT_TXT = 'output_text.txt'


def count_connections(data, print_connections=False):
    num_connections = np.zeros(len(data), dtype=int)
    for column in data.iloc[:, 2:]:
        for value in data[column]:
            num_connections[value - 1] += 1

    if print_connections:
        for index, value in data['שם מלא'].items():
            print(value, ': ', num_connections[data['מספר שלי'][index] - 1])

    return num_connections


def count_connections_by_name(data, connections, name):
    name_idx = data.index[data['שם מלא'] == name].tolist()
    if len(name_idx) == 0:
        return -1
    cadet_idx = data['מספר שלי'][name_idx[0]]
    return connections[cadet_idx - 1]


def highlight_cells(value):
    if isinstance(value, str):
        return ''
    if value <= LOW_THRESH:
        return 'color: red'
    elif value >= HIGH_THRESH:
        return 'color: green'
    else:
        return ''


def highlight_diff(s):
    result = [''] * len(s.values)
    for i in range(2, len(s.values)):
        if s.values[i] - s.values[i-1] >= DIFF_THRESH:
            result[i] = 'background-color: lightblue'
        elif s.values[i] - s.values[i-1] <= -DIFF_THRESH:
            result[i] = 'background-color: lightyellow'
    return result


def get_popularity_over_time(dfs, save_to_file=False):
    results = dict()
    connections = []
    summary_str = ""
    for i in range(len(dfs)):
        connections.append(count_connections(dfs[i]))
    for index, value in dfs[-1]['שם מלא'].items():
        num_connections = [count_connections_by_name(dfs[i], connections[i], value) for i in range(len(dfs))]
        results[value] = num_connections
        if num_connections[-1] <= LOW_THRESH:
            summary_str += value + ' has only ' + str(num_connections[-1]) + ' connections\n'
        if num_connections[-1] - num_connections[0] >= DIFF_THRESH:
            summary_str += value + ' has ' + str(num_connections[-1]) + ' more connections\n'
        elif num_connections[-1] - num_connections[0] <= -DIFF_THRESH:
            summary_str += value + ' has ' + str(num_connections[0] - num_connections[-1]) + ' less connections\n'

    if save_to_file:
        res_file = pd.DataFrame()
        # Add a 'name' column to the DataFrame with the dictionary keys
        res_file['Name'] = results.keys()
        for i in range(len(dfs)):
            res_file['Semester ' + str(i + 1)] = [value[i] for value in results.values()]

        styled_df = res_file.style.applymap(highlight_cells).apply(highlight_diff, axis=1)
        styled_df.to_excel(OUTPUT_EXCEL, index=False)

        with open(OUTPUT_TXT, 'w', encoding='utf-8') as file:
            file.write(summary_str)

        print(res_file)


if __name__ == '__main__':
    load_file_to_df = True
    file_paths = ["C:/Users/t9155275/Desktop/sociometry_md_last_sem/md_semester_a.xlsx",
                  "C:/Users/t9155275/Desktop/sociometry_md_last_sem/socoio_md_semester_c.xlsx"]

    dfs = [pd.read_excel(file_path, header=None, engine="openpyxl") for file_path in file_paths]
    for i in range(len(dfs)):
        dfs[i].columns = dfs[i].iloc[0]
        dfs[i] = dfs[i].drop(dfs[i].index[0])
        dfs[i]['שם מלא'] = dfs[i]['שם מלא'].str.strip()

    get_popularity_over_time(dfs, True)
