# AssesmentTalpiotCodes

Hi, there are 4 events which you have code for: Miun feedbacks, Socio, Sagab-Sagaz feedback, Sociogram

The code for Socio+Sagab-Sagaz is pretty good. The codes for Miun feedbacks and Sociogram are less generic but working.

## Socio + Sagab-Sagaz

In the directory of Socio_and_Sagabz their is a code for this 2 events.
Pay attention for all of the constants in the code, those constants determine the run_kind.
You need the socio be in the right formats as in the format directory.

Input: Excels of Socio or big responses excel for sagaz-sagab feedback (The excel formats are really important)
(Input example at "Formats/Socio_input_directory_format" and "Formats/Sagaz-Sagab-input-format.xlsx" - this file is created by google forms)
Output: Directory with results

There are options for:
1. remove people with less than filter_by_num_of_answers.
2. adding old stats (from last year) - these will add bar to word graphs from last year.

## Sociogram

For the sociogram you need to run the notebook step after step, at the end you will get graphs that you can put in presentation.
The results are from google froms sheet (can find a format in formats)
There is example presentation for the graph results.

# Main Sociogram

This code produces a slide for each cadet, a slide for each group it identifies, and histograms that show the connectivity
of different subgraphs. 
- The PrivateSlides class produces a slide that contains a graph of his connections, and data regarding the number of
unidirectional and bidirectional connections he has.
- The CliqueSlides class starts by identifying all cliques of a certain size (with no overlapping nodes), and then iterates
over all remaining nodes. For each node, it calculates the number of connections it has to each clique. If the clique it
has the most connections to has more than a specified threshold of all connections of that node (i.e 80% of a node's
connections are related to some clique), it assigns the node to this clique.
It then generates a slide for each group coloring the original clique in blue, and the additional nodes in light blue.
It states the dominant features of the nodes in the group as well as the most popular cadets in the group (determined by
the largest in-degree of nodes in the subgraph that is the group).
- sociogram.py generates histograms based on the categories defined in the datasheet. It the gives each category a score
of connectivity based on the actual amount of inner-connections in the subgroup, divided by the expected number of connections
(based on the size of the subgroup). This means that scores above 1 indicate a more strongly connected group and vice versa.

In order to use this code, provide the necessary file paths to the saved tables by chronological order, and the desired
semester index (where zero is the semester that first appears in the file_paths list).

## Miun feedbacks

This is code for feedback files for miun.
You can run the codes and see the results. for each talpion their will be a word with it's results.
(The excel file is censored - wothout the names)