import pandas as pd
from tqdm import tqdm

from Sociogram import sociogram
from private_slides import *
from clique_slides import *

FILE_PATHS = [r"C:\Work\Haaraha\assesmenttalpiotcodes-main\Sociogram\test.xlsx",
              r"C:\Work\Haaraha\assesmenttalpiotcodes-main\Sociogram\test.xlsx"
              ]  # the file paths sorted chronologically
# FILE_PATHS = ['data_example/Sociogram_Results.xlsx']
FILE_OUT = r"C:\Work\Haaraha\assesmenttalpiotcodes-main\Sociogram"  # the output folder, must be an existing folder!

SEMESTER_INDEX = 1  # semester index, based on the index in the file_paths list

IDENTIFIER = 'מס תז'
NAME = 'שם מלא'  # name column title
NUM = 'מספר שלי'  # number column title
FEATURES_FOR_HISTOGRAM = ['מחלקה', 'מין', 'מגזר', 'רקע אקדמי',"מחלקה בטירונות", "מסדרון"]  # features to characterize groups by
FEATURES_FOR_PRIVATE_SLIDES = ["מסדרון",'מחלקה', 'מין', 'מגזר', 'רקע אקדמי']  # features to characterize groups by
# FEATURES = []
CHOICES = [f'בחירה {i}' for i in range(1, 7)]  # choices column titles
CHOICES_LIST = 'רשימת בחירות'  # for inner use

PLOT_HISTOGRAMS = False  # whether to plot histograms or not

"""
Currently shown stats:
1. 'In-Degree' (no list)
2. 'Previous In-Degree' (no list)
3. 'Centrality' (no list)
4. 'Bidirectional Edges' (list)
5. 'Unidirectional Edges from Node' (list)
6. 'Unidirectional Edges to Node' (list)
7. 'Connections Formed' (list)
8. 'Connections Deformed' (list)
9. 'Pion Friends' (list)
"""
SHOW_LIST = [True, True, True, True, True, True]  # whether to show the list in the graph description or not for every stat
NO_LIST_STATS = 3  # number of stats which do not contain lists
LIST_STATS = 6  # number of stats which contain lists

MIN_CLIQUE_SIZE = 3
MAX_CLIQUE_SIZE = 15
MAX_NUM_CLIQUES = 6
ASSIGN_NODE_THRESHOLD = 0.3  # if a node has at least this part of edges connected to a certain clique, it will be assigned to that clique
ASSIGNING_ITERATIONS = 3  # number of iterations to perform assigning remaining nodes to cliques
FEATURE_THRESHOLD = 0.5  # part of the group that needs to have this feature in order to consider it a dominant feature
NUM_DOMINANT_NODES = 3  # the number of dominant nodes to display in the title of the group graph


def create_graph(dfs: List[pd.DataFrame], num_to_identifier: List[Dict[int, str]], index: int) -> nx.DiGraph:
    """
    Create a directed graph based on the input data for a specific semester.

    :param dfs: List of DataFrames.
    :param num_to_identifier: Dictionary mapping numbers to cadet identifier (usually id).
    :param index: Index of the semester to create the graph for.
    :return: A directed graph representing social connections for the specified semester.
    """
    df = dfs[index]
    g = nx.DiGraph()  # creating a directed graph
    g.add_nodes_from(df[IDENTIFIER].tolist())  # adding the list of names as nodes to the graph
    for idx, row in df.iterrows():  # iterating over cadets
        u = row[IDENTIFIER]

        # if row[CHOICES_LIST] has nan values then remove them (if someone didnt put 6 friends)
        row[CHOICES_LIST] = [x for x in row[CHOICES_LIST] if str(x) != 'nan']

        vs = [num_to_identifier[index][x] for x in row[CHOICES_LIST]]  # per cadet, this is the list of his/her chosen friends
        for v in vs:
            g.add_edge(u, v)  # iterating over the chosen friends and adding the edges
    return g


if __name__ == '__main__':
    if SEMESTER_INDEX >= len(FILE_PATHS):
        raise Exception('Invalid semester index!')

    dfs = [pd.read_excel(FILE_PATHS[i]) for i in range(len(FILE_PATHS))]

    for i in range(len(dfs)):
        dfs[i][NAME] = dfs[i][NAME].str.strip()
        dfs[i][CHOICES_LIST] = dfs[i].loc[:, CHOICES].apply(lambda r: r.tolist(), axis=1)  # transforming the list of columns into a column containing lists of the choices per cader

    num_to_identifier = [{row[NUM]: row[IDENTIFIER] for _, row in dfs[i].iterrows()} for i in range(len(dfs))]  # dictionary from <cadet num> to <cadet name>
    num_to_choices = [{row[NUM]: row[CHOICES] for _, row in dfs[i].iterrows()} for i in range(len(dfs))]
    identifier_to_num = [{row[IDENTIFIER]: row[NUM] for _, row in dfs[i].iterrows()} for i in range(len(dfs))]
    identifier_to_name = {row[IDENTIFIER]: row[NAME] for _, row in dfs[0].iterrows()}

    try:
        identifier_to_features = [{row[IDENTIFIER]: [row[c] for c in FEATURES_FOR_PRIVATE_SLIDES] for _, row in dfs[i].iterrows()} for i in range(len(dfs))]
    except Exception as e:
        print("Required features not found in table!")
        identifier_to_features = [{row[IDENTIFIER]: [] for _, row in dfs[i].iterrows()} for i in range(len(dfs))]

    graphs = [create_graph(dfs, num_to_identifier, i) for i in range(len(dfs))]

    # plot private cadet slides
    print("plotting slides for each cadet...")
    privateSlides = PrivateSlides(dfs, graphs, num_to_identifier, identifier_to_num, identifier_to_name, num_to_choices)
    for identifier in tqdm(graphs[SEMESTER_INDEX].nodes()):
        privateSlides.plot_sociogram(identifier)

    # plot group slides
    print("plotting slides for each group...")
    cliqueSlides = CliqueSlides(graphs[SEMESTER_INDEX], identifier_to_features[SEMESTER_INDEX], identifier_to_name)
    non_overlapping_cliques, remaining_vertices = cliqueSlides.find_cliques_with_limit()
    second_order_nodes = cliqueSlides.assign_remaining_nodes_to_cliques(non_overlapping_cliques, remaining_vertices)
    for i in tqdm(range(len(non_overlapping_cliques))):
        cliqueSlides.draw_graph_with_second_order_nodes(non_overlapping_cliques, second_order_nodes, i)

    # plot histograms
    if PLOT_HISTOGRAMS:
        try:
            print("plotting histograms for each category...")
            sociogram.plot_histogram(FILE_PATHS[SEMESTER_INDEX])
        except Exception as e:
            print(e)




