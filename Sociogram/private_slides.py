import os
import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
from typing import List, Dict

from Sociogram import main_sociogram


class PrivateSlides:

    def __init__(self, dfs: List[pd.DataFrame], graphs: List[nx.DiGraph], num_to_identifier: List[Dict[int, str]],
                 identifier_to_num: List[Dict[str, int]], identifier_to_name: Dict[str, str],
                 num_to_choices: List[Dict[int, pd.DataFrame]]):
        """
        Initialize PrivateSlides instance.

        :param dfs: List of DataFrames.
        :param graphs: List of directed graphs.
        :param num_to_identifier: Dictionary mapping numbers to cadet names.
        :param identifier_to_num: Dictionary mapping cadet names to numbers.
        :param num_to_choices: Dictionary mapping numbers to DataFrame of choices.
        :return: None
        """
        self.dfs = dfs
        self.num_to_identifier = num_to_identifier
        self.identifier_to_num = identifier_to_num
        self.identifier_to_name = identifier_to_name
        self.num_to_choices = num_to_choices
        self.graphs = graphs

        self.semester_idx = main_sociogram.SEMESTER_INDEX
        self.identifier_col = main_sociogram.IDENTIFIER
        self.choices_list_col = main_sociogram.CHOICES_LIST
        self.no_list_stats = main_sociogram.NO_LIST_STATS
        self.list_stats = main_sociogram.LIST_STATS
        self.show_list = main_sociogram.SHOW_LIST

        self.pions = list(set(dfs[0].loc[:, self.identifier_col]) - set(dfs[-1].loc[:, self.identifier_col]))
        self.bigraphs = [PrivateSlides.filter_bidirectional(g) for g in graphs]
        self.stats_df = self.calculate_node_features()

    def print_description(self, data: pd.Series) -> str:
        description = ""
        total_stats = self.no_list_stats + self.list_stats
        for i in range(1, 1 + self.no_list_stats):
            if self.semester_idx == 0 and i == 2:
                continue
            description += data.index[i] + ": " + str(data.iloc[i]) + " (" + str(data.iloc[i + total_stats]) + ")\n\n"
        for i in range(1 + self.no_list_stats, (1 if self.semester_idx > 0 else -2) + self.no_list_stats + self.list_stats):
            description += data.index[i] + ": " + str(len(data.iloc[i])) + " (" + str(data.iloc[i + total_stats]) + ")"
            if self.show_list[i - 1 - self.no_list_stats] and len(data.iloc[i]) > 0:
                description += '\n' + ', '.join([self.identifier_to_name[n][::-1] for n in data.iloc[i]])
            description += "\n\n"
        return description

    def calculate_node_features(self) -> pd.DataFrame:
        """
        Calculate various node features for a given semester.

        :return: A DataFrame containing calculated node features.
        """
        # Initialize lists to store feature values for each node
        g = self.graphs[self.semester_idx]
        nodes = list(g.nodes())
        in_degrees = [g.in_degree(node) for node in nodes]
        centrality = [nx.closeness_centrality(g, node) for node in nodes]

        # these are the absolute stats per-cadet in the graph
        bidirectional_edges = [[v for u, v in g.out_edges(node) if g.has_edge(v, node)] for node in nodes]
        unidirectional_from_nodes = [[v for u, v in g.out_edges(node) if not g.has_edge(v, node)] for node in nodes]
        unidirectional_to_nodes = [[u for u, _ in g.in_edges(node) if not g.has_edge(node, u)] for node in nodes]

        if self.semester_idx > 0:
            # kind of voodo magic if there is problem here talk with Iftah Farkash if hes still in Talpiot
            prev_in_degrees = [self.graphs[self.semester_idx-1].in_degree(node) for node in nodes]
            connections_formed = [
                [u for u, v in self.bigraphs[self.semester_idx].in_edges(node)
                 if not self.graphs[self.semester_idx-1].has_edge(u, v)
                 and not self.graphs[self.semester_idx-1].has_edge(v, u)] for node in nodes]
            connections_deformed = [[u for u, v in self.bigraphs[self.semester_idx - 1].in_edges(node)
                                     if u not in self.pions and not g.has_edge(u, v)
                                     and not g.has_edge(v, u)] for node in nodes]
        else:
            prev_in_degrees, connections_formed, connections_deformed = [0] * len(nodes), [[]] * len(nodes), [[]] * len(nodes)

        pion_friends = [[p for p in self.pions if self.are_good_friends(node, p)] for node in nodes]
        # Create a DataFrame
        data = {
            self.identifier_col: nodes,
            'In-Degree': in_degrees,
            'Previous In-Degree': prev_in_degrees,
            'Centrality': centrality,
            'Bidirectional Edges': bidirectional_edges,
            'Unidirectional Edges from Node': unidirectional_from_nodes,
            'Unidirectional Edges to Node': unidirectional_to_nodes,
            'Connections Formed': connections_formed,
            'Connections Deformed': connections_deformed,
            'Pion Friends': pion_friends
        }
        df = pd.DataFrame(data)
        names_col = df[self.identifier_col].map(self.identifier_to_name)
        df_copy = df.copy()
        df_copy.insert(1, 'Name', names_col)
        df_copy.to_excel(main_sociogram.FILE_OUT + '/table.xlsx', index=False)

        # Calculate percentiles among all nodes for each feature
        for i in range(1, len(df.columns)):
            feature = df.columns[i]
            if i < self.no_list_stats + 1:
                percentile_values = df[feature].rank(pct=True)
            else:
                percentile_values = df[feature].apply(lambda x: len(x)).rank(pct=True)
            df[f'{feature} Percentile'] = percentile_values.round(2)

        return df

    def plot_sociogram(self, identifier: str) -> None:
        """
        Generate and save a sociogram for a specific cadet's name.

        :param identifier: Identifier of the cadet.
        :return: None
        """
        # creating the figure and adjusting sizes
        g = self.graphs[self.semester_idx]
        name = self.identifier_to_name[identifier].strip()
        gs = GridSpec(1, 2, width_ratios=[3, 1])  # Adjust the height ratios as needed
        fig = plt.figure(figsize=(12, 8))
        ax = [fig.add_subplot(gs[0]), fig.add_subplot(gs[1])]

        # obtaining the list of nodes which are adjacent to the given cadet
        nodes = [u for u in g.nodes() if (g.has_edge(u, identifier) | g.has_edge(identifier, u))]
        nodes.append(identifier)
        if self.semester_idx > 0:
            ex_friends = [u for u, v in self.bigraphs[self.semester_idx-1].in_edges(identifier) if u not in self.pions and not g.has_edge(u, v) and not g.has_edge(v, u)]
        else:
            ex_friends = []
        nodes.extend(ex_friends)

        # drawing the subgraph containing only the above nodes
        self.help_draw(g.subgraph(nodes), center=identifier, ex_friends=ex_friends, ax=ax[0])

        # adding text, title, and saving the figure
        ax[1].text(0, .4, self.print_description(self.stats_df[self.stats_df[self.identifier_col] == identifier].squeeze()))
        ax[1].axis('off')
        fig.suptitle(f'{name[::-1]}')
        # fig.suptitle(f'{name} - סמסטר {self.semester_idx + 1}'[::-1])

        output = main_sociogram.FILE_OUT + '/cadets'
        if not os.path.exists(output):
            os.makedirs(output)
        file_path = os.path.join(output, f'{name}.png')
        plt.savefig(file_path)

    def are_good_friends(self, name1: str, name2: str) -> bool:
        """
        Helper function to determine if two cadets are good friends.

        :param name1: Name of the first cadet.
        :param name2: Name of the second cadet.
        :return: True if the cadets are good friends, False otherwise.
        """
        for i in range(len(self.dfs)):
            if name1 not in self.identifier_to_num[i] or name2 not in self.identifier_to_num[i]:
                return False
            num1 = self.identifier_to_num[i][name1]
            num2 = self.identifier_to_num[i][name2]
            if num2 in self.num_to_choices[i][num1].tolist() and num1 in self.num_to_choices[i][num2].tolist():
                return True
        return False

    def help_draw(self, g: nx.DiGraph, center: str, ex_friends: List[str], ax: plt.Axes = None) -> None:
        """
        Helper function to draw a sociogram based on the provided graph and additional parameters.

        :param g: Directed graph.
        :param center: The central node around which the sociogram is drawn.
        :param ex_friends: List of nodes representing friends who are excluded.
        :param ax: Matplotlib axis to plot on.
        :return: None
        """
        # Reverse the node labels (Hebrew bug...)
        try:
            flipped_node_names = {node: self.identifier_to_name[node][::-1].replace(' ', '\n') for node in g.nodes()}
        except:
            print(g.nodes())
        # Compute Kamada-Kawai layout as the initial position - this is a method to embed the graph in R^2
        pos = nx.kamada_kawai_layout(g)

        # Run spring layout for a specific number of iterations - this is a tweak i added to make the nodes nicer-looking in terms of positions
        pos = nx.spring_layout(g, pos=pos, iterations=100, k=4)

        # Adjust x-coordinate to place labels to the left of each node
        label_pos = {k: (v[0] - 0.07, v[1]) for k, v in pos.items()}

        connections_formed = list(self.stats_df[self.stats_df[self.identifier_col] == center]['Connections Formed'])[0]

        if ax == None:
            fig, ax = plt.subplots(figsize=(12, 8))

        # Draw custom node labels
        nx.draw_networkx_labels(g, label_pos, labels=flipped_node_names, font_size=12, font_color='black', ax=ax)

        # Draw nodes and edges with custom arrow and edge settings
        edges = g.edges()
        bidirectional_edges = [(u, v) for u, v in edges if (v, u) in edges]
        broken_edges = [(u, center) for u in ex_friends]
        formed_edges = [(u, center) for u in connections_formed] + [(center, u) for u in connections_formed]

        unidirectional_edges = list(set(edges) - set(bidirectional_edges))
        bidirectional_edges = list(set(bidirectional_edges) - set(formed_edges))

        # drawing the graph
        node_color = ['k' if node == center else ('red' if node in ex_friends else 'grey') for node in g.nodes()]
        nx.draw(g, pos=pos, with_labels=False, edgelist=unidirectional_edges, edge_color='green', arrowsize=20,
                alpha=0.5,
                ax=ax)
        nx.draw(g, node_color=node_color, pos=pos, with_labels=False, edgelist=bidirectional_edges, edge_color='green',
                width=2, alpha=0.8, arrows=False, ax=ax)
        nx.draw_networkx_edges(g, pos=pos, edgelist=broken_edges, edge_color='red', alpha=0.8, arrows=False, ax=ax)
        nx.draw_networkx_edges(g, pos=pos, edgelist=formed_edges, edge_color='blue', alpha=0.8, arrows=False, ax=ax)

    @staticmethod
    def filter_bidirectional(g: nx.DiGraph) -> nx.DiGraph:
        """
        Filter bidirectional edges from a directed graph.

        :param g: Directed graph.
        :return: A new directed graph containing only bidirectional edges.
        """
        bidirectional_graph = nx.DiGraph()
        bidirectional_graph.add_nodes_from(g.nodes())

        bidirectional_edges = [(u, v) for u, v in g.edges() if g.has_edge(v, u)]
        bidirectional_graph.add_edges_from(bidirectional_edges)

        return bidirectional_graph

    @staticmethod
    def plot_node_features_table(df: pd.DataFrame, node_name: str, ax: plt.Axes = None) -> None:
        """
        Plot a table of node features for a specific node.

        :param df: DataFrame containing node features.
        :param node_name: Name of the node to plot features for.
        :param ax: Matplotlib axis to plot on.
        :return: None
        """
        # Filter the DataFrame to get the relevant row for the specified node
        node_row = df[df['Node'] == node_name]

        # Create a table with the extracted row
        if ax is None:
            fig, ax = plt.subplots(figsize=(8, 2))

        ax.axis('off')  # Turn off axis for better visualization
        ax.table(cellText=node_row.values[:, 1:], colLabels=node_row.columns[1:], cellLoc='center', loc='center')

        # Set the title
        ax.set_title(f"Node Features for Node {node_name}")