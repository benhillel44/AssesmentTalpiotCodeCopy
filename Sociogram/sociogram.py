import os

import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
import networkx as nx

from Sociogram import main_sociogram


def get_count_dict(df_to_count, nodes_to_count=None):
    """
    get dict {node: num of connections to}
    :param df_to_count: df to count all connections of a subgroup
    :param edges_columns: the columns of connections
    :param nodes_to_count: count connections only for specific subgroup
    :return:
    """
    edges_columns = main_sociogram.CHOICES
    count_dict = dict()
    for edge in edges_columns:
        for node in list(df_to_count[edge].values):
            # the case want to count only connections with specific subgraph
            if (nodes_to_count is not None) and (node not in nodes_to_count):
                continue
            if node in count_dict:
                count_dict[node] += 1
            else:
                count_dict[node] = 1
    return count_dict


def get_group_connection_value(df, category, group_name):
    """
    get the connectivity value
    :param df: relevant df
    :param category: the column name. i.e. "gender"
    :param group_name: the value of the column. i.e. "female"
    :return:
    """
    size = len(df)
    subgroup_df = df[df[category] == group_name]
    subgroup_size = len(subgroup_df)

    # get expected inner connections - num_of_connections_to_group * (size_group/total_size)
    all_connections_to_subgroup = get_count_dict(df_to_count=df, nodes_to_count=list(subgroup_df[main_sociogram.NUM].unique()))
    num_all_connections = sum([all_connections_to_subgroup[k] for k in all_connections_to_subgroup.keys()])
    expected_inner_connections = num_all_connections * (subgroup_size/size)

    # get inner connections
    inner_connections_to_subgroup = get_count_dict(df_to_count=subgroup_df, nodes_to_count=list(subgroup_df[main_sociogram.NUM].unique()))
    num_inner_connections = sum([inner_connections_to_subgroup[k] for k in inner_connections_to_subgroup.keys()])
    inner_connections = num_inner_connections

    return inner_connections/expected_inner_connections


def plot_hists_by_category(df):
    """
    plot histogram of connectivity value by category
    :param df: relevant dataframe
    :return:
    """
    categories = main_sociogram.FEATURES_FOR_HISTOGRAM
    for category in categories:
        group_names = []
        connectivity_values = []
        for group_name in sorted(df[category].unique()):
            conn_value = get_group_connection_value(df, category, group_name)
            connectivity_values.append(conn_value)
            group_names.append(str(group_name))

        fig = px.histogram(x=group_names, y=connectivity_values)
        fig.add_hline(y=1)
        fig["layout"]["yaxis"]["title"]["text"] = category
        fig.show()


def plot_histogram(file_path):
    # load_file_to_df = True
    # file_path = r"data_example/סוציוגרם מד - סמסטר א.xlsx"

    df = pd.read_excel(file_path, header=None, engine="openpyxl")
    df.columns = df.iloc[0]
    df = df.drop(df.index[0])
    df["מספר חדר"] = df["מספר חדר"].astype(float)  # some format issues

    plot_connectivity_hists_by_category = True
    if plot_connectivity_hists_by_category:
        plot_hists_by_category(df)

    # print_number_connections = True
    # if print_number_connections:
    #     count_dict = get_count_dict(df)
    #     for key in count_dict:
    #         print(f"{df[df['שם מלא - מספר'] == key]['שם מלא'].iloc[0]}: {count_dict[key]}")