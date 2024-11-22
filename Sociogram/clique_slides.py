import os

import networkx as nx
from matplotlib import pyplot as plt
from matplotlib.gridspec import GridSpec

from typing import List, Dict, Tuple, Any, Set, Union

from Sociogram import main_sociogram


class CliqueSlides:

    def __init__(self, graph: nx.DiGraph, identifier_to_features: Dict[str, List[str]],
                 identifier_to_name: Dict[str, str]):
        """
        Initialize CliqueSlides instance.

        :param graph: The directed graph.
        :param identifier_to_features: Dictionary mapping cadet names to a list of features.
        :return: None
        """
        self.graph = graph
        self.min_clique_size = main_sociogram.MIN_CLIQUE_SIZE
        self.max_clique_size = main_sociogram.MAX_CLIQUE_SIZE
        self.max_num_cliques = main_sociogram.MAX_NUM_CLIQUES
        self.num_dominant_nodes = main_sociogram.NUM_DOMINANT_NODES
        self.assign_thresh = main_sociogram.ASSIGN_NODE_THRESHOLD
        self.num_iterations = main_sociogram.ASSIGNING_ITERATIONS
        self.feature_thresh = main_sociogram.FEATURE_THRESHOLD
        self.identifier_to_features = identifier_to_features
        self.identifier_to_name = identifier_to_name

    def find_cliques_with_limit(self) -> tuple[list[Union[list[Any] , Any]], set[Any]]:
        """
        Find cliques in the graph. The cliques are found usign the built-in find_cliques method of the networkx library
        and then sorted by their size (and by lexical order, for the algorithm to be deterministic). The cliques are
        then iterated over, and overlapping nodes between cliques are removed so that the cliques have unique nodes.

        :return: A tuple containing non-overlapping cliques and remaining vertices.
        """
        graph = self.graph.to_undirected()

        # Find all cliques in the graph
        all_cliques = sorted(list(nx.find_cliques(graph)), key=len, reverse=True)
        # Sort each clique internally by lexical order
        all_cliques = [sorted(clique) for clique in all_cliques]
        # Sort all cliques by their length and lexical order
        all_cliques = sorted(all_cliques, key=lambda x: (len(x), x))

        # Select non-overlapping and potentially overlapping cliques based on the threshold
        non_overlapping_cliques = []
        overlapping_cliques = []
        nodes_covered = set()

        for clique in all_cliques:
            if not any(node in nodes_covered for node in clique) and len(clique) > self.min_clique_size:
                non_overlapping_cliques.append(clique)
                nodes_covered.update(clique)

                # Check if the desired number of cliques is reached
                if self.max_num_cliques != -1 and len(non_overlapping_cliques) >= self.max_num_cliques:
                    break

            elif len(clique) > self.min_clique_size:
                overlapping_cliques.append(clique)

        # Check for overlapping nodes in the second iteration only if the maximum number of cliques is not reached
        if self.max_num_cliques == -1 or len(non_overlapping_cliques) < self.max_num_cliques:
            for overlapping_clique in overlapping_cliques:
                overlapping_set = set(overlapping_clique)
                overlapping_nodes = overlapping_set.intersection(nodes_covered)

                # If removing overlapping nodes still gives a clique larger than the threshold, move to non-overlapping
                if len(overlapping_nodes) > 0 and len(overlapping_clique) - len(overlapping_nodes) > self.min_clique_size:
                    non_overlapping_cliques.append(list(overlapping_set - overlapping_nodes))
                    nodes_covered.update(overlapping_set - overlapping_nodes)

                    # Check if the desired number of cliques is reached
                    if self.max_num_cliques != -1 and len(non_overlapping_cliques) >= self.max_num_cliques:
                        break

        remaining_vertices = set(graph.nodes()) - nodes_covered
        return non_overlapping_cliques, remaining_vertices

    def assign_remaining_nodes_to_cliques(self, non_overlapping_cliques: List[List[str]], remaining_nodes: set[str]) -> List[List[str]]:
        """
        Assign remaining nodes to cliques. Each node counts the number of connections it has to each group, and then
        finds the group it is closest to. If the part of connections the node has to this group from all of its
        connections is larger than the specified threshold, the node is assigned to this group. This process is performed
        several iterations.

        :param non_overlapping_cliques: Non-overlapping cliques.
        :param remaining_nodes: Remaining nodes.
        :return: List of second order nodes.
        """
        second_order_nodes = [[] for _ in range(len(non_overlapping_cliques))]

        for _ in range(self.num_iterations):
            nodes_to_remove = []
            average_clique_size = sum(len(non_overlapping_cliques[k] + second_order_nodes[k]) for k in range(len(non_overlapping_cliques))) / len(non_overlapping_cliques)
            remaining_nodes = sorted(remaining_nodes)

            for node in remaining_nodes:
                max_edges = 0
                closest_clique_index = -1
                node_edges_count = len(self.graph.in_edges(node)) + len(self.graph.out_edges(node))

                for i in range(len(non_overlapping_cliques)):
                    clique = non_overlapping_cliques[i] + second_order_nodes[i]
                    if len(clique) > self.max_clique_size:
                        continue
                    # Count the number of edges between the node and members of the clique
                    edges_to_clique = sum(1 for member in clique if (node, member) in self.graph.edges() or (member, node) in self.graph.edges())
                    edges_from_clique = sum(1 for member in clique if (member, node) in self.graph.edges() or (node, member) in self.graph.edges())
                    total_edges = (edges_to_clique + edges_from_clique) * average_clique_size / len(clique)

                    # Update the closest clique if the current clique has more edges
                    if total_edges > max_edges:
                        max_edges = total_edges
                        closest_clique_index = i

                    # print(node, total_edges, node_edges_count, clique)
                    # if total_edges > threshold * node_edges_count and node not in second_order_nodes[i]:
                    #     second_order_nodes[i].append(node)

                # Add the node to the closest clique in second_order_nodes
                if closest_clique_index != -1 and max_edges > self.assign_thresh * node_edges_count:
                    # print(node, non_overlapping_cliques[closest_clique_index], _, closest_clique_index)
                    second_order_nodes[closest_clique_index].append(node)
                    nodes_to_remove.append(node)

            remaining_nodes = list(set(remaining_nodes) - set(nodes_to_remove))

        remaining_names = [self.identifier_to_name[node] for node in remaining_nodes]
        to_write = 'Nodes not assigned to any group: ' + ', '.join(sorted(remaining_names))
        print(to_write)
        with open(main_sociogram.FILE_OUT + '/Unassigned Nodes.txt', 'w', encoding='utf-8') as file:
            file.write(to_write)
        return second_order_nodes

    def draw_graph_with_second_order_nodes(self, cliques: List[List[str]], second_order_nodes: List[List[str]], index: int) -> None:
        """
        Draw the graph with second-order nodes.

        :param cliques: List of cliques.
        :param second_order_nodes: List of second order nodes.
        :param index: Index of the clique.
        :return: None
        """
        # Add original nodes of the clique
        gs = GridSpec(1, 2, width_ratios=[3, 1])  # Adjust the height ratios as needed
        fig = plt.figure(figsize=(12, 8))
        g = self.graph.subgraph(cliques[index] + second_order_nodes[index])

        ax = [fig.add_subplot(gs[0]), fig.add_subplot(gs[1])]
        ax[1].text(0, .4, self.characterize_group(cliques[index] + second_order_nodes[index]))
        ax[1].axis('off')
        # fig.suptitle(f'קבוצה מספר {index+1}'[::-1])
        fig.suptitle(f'קבוצה של {self.get_dominant_nodes(g)}'[::-1])

        flipped_node_names = {node: self.identifier_to_name[node][::-1].replace(' ', '\n') for node in g.nodes()}
        pos = nx.kamada_kawai_layout(g)
        pos = nx.spring_layout(g, pos=pos, iterations=100, k=4)
        label_pos = {k: (v[0] - 0.07, v[1]) for k, v in pos.items()}
        # nx.draw_networkx_labels(g, label_pos, labels=flipped_node_names, font_size=12, font_color='black', ax=ax)
        node_colors = ['blue' if node in cliques[index] else 'deepskyblue' for node in g.nodes()]

        # Draw the graph
        nx.draw(g, pos, with_labels=True, labels=flipped_node_names, node_color=node_colors, font_color='black', node_size=700, ax=ax[0])

        output = main_sociogram.FILE_OUT + '/groups'
        if not os.path.exists(output):
            os.makedirs(output)
        file_path = os.path.join(output, f'group {index+1}.png')
        plt.savefig(file_path)

    def characterize_group(self, nodes: List[str]) -> str:
        """
        Characterize a group based on features.

        :param nodes: List of nodes in the group.
        :return: Group characteristics.
        """
        if not nodes:
            return ""

        features_count = dict()
        for features in self.identifier_to_features.values():
            for f in features:
                if f not in features_count:
                    features_count[f] = [1, 0]
                else:
                    features_count[f][0] += 1
        total_nodes = len(nodes)

        # Count the number of nodes with each feature
        for node in nodes:
            feature = self.identifier_to_features.get(node)
            if feature:
                for f in feature:
                    features_count[f][1] += 1

        # Print characteristics based on the criteria
        description = ""
        for feature, count in features_count.items():
            if count[1] / total_nodes >= max(self.feature_thresh, min(0.99, 1.1 * count[0] / len(self.identifier_to_features.keys()))):
                description += f"{round(100 * count[1] / total_nodes, 1)}% of the group are {feature[::-1]}\n"

            if count[0] > 0 and count[1] / count[0] > self.feature_thresh:
                description += f"{round(100 * count[1] / count[0], 1)}% of all nodes with {feature[::-1]}\nfeature are in this group\n"

        if description == "":
            return "No dominant features in this group"
        else:
            return "Dominant features of the group:\n" + description

    def get_dominant_nodes(self, graph: nx.DiGraph) -> str:
        """
        Get the dominant nodes in the graph based on in-degree in the subgraph that is the group. In other words, the
        dominant nodes are those that have the most arrows going in, from within the specified group.
        Only the top most <num_dominant_nodes> are taken into account in the final graph title.

        :param graph: The directed graph.
        :return: Dominant nodes.
        """
        sorted_nodes = sorted(graph.nodes, key=lambda x: graph.in_degree(x), reverse=True)
        top_nodes = sorted_nodes[:self.num_dominant_nodes]
        top_nodes = [self.identifier_to_name[node] for node in top_nodes]
        return ', '.join(top_nodes)
