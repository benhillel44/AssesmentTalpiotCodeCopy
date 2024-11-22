import itertools

import networkx as nx
import numpy as np
import pandas as pd


def sort_data(dfs):
    connections = []
    dicts = []
    names = []
    for i in range(len(dfs)):
        df_sorted = dfs[i].sort_values(by='מספר שלי')
        np_array = df_sorted.iloc[:, 2:].values
        np_array = np_array - 1
        names_list = list(df_sorted['שם מלא'])
        name_index_dict = dfs[i].set_index('שם מלא')['מספר שלי'].to_dict()
        for name in name_index_dict:
            name_index_dict[name] -= 1
        connections.append(np_array)
        names.append(names_list)
        dicts.append(name_index_dict)
    return names, dicts, connections


def get_date_from_filepaths(file_paths):
    dfs = [pd.read_excel(file_path, header=None, engine="openpyxl") for file_path in file_paths]
    for i in range(len(dfs)):
        dfs[i].columns = dfs[i].iloc[0]
        dfs[i] = dfs[i].drop(dfs[i].index[0])
        dfs[i]['שם מלא'] = dfs[i]['שם מלא'].str.strip()
    return dfs


def symmetrize_matrix(A, degree=2):
    A = 10*A
    if degree >= 2:
        A += A@A + A@A.T
    if degree >= 3:
        A += A@A@A + A@A@A.T + A@A.T@A + A.T@A@A
    A = A + A.T
    for i in range(len(A)):
        A[i][i] = 0
    return A


def girvan_newman_method(g, k):
    comp = nx.community.girvan_newman(g)
    communities = ([i for i in range(len(g))], [])
    for step in itertools.islice(comp, k - 1):
        communities = tuple(sorted(c) for c in step)
    return communities


def louvain_method(g):
    comp = nx.community.louvain_communities(g)
    communities = tuple(sorted(c) for c in comp)
    return communities