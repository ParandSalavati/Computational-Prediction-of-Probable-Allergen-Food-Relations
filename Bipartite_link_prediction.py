import csv
import pandas as pd
import networkx as nx
from networkx.algorithms import bipartite

# Load data from an Excel file
df = pd.read_excel('Top10Allergens.xlsx')

# Create a bipartite graph from the DataFrame
B = nx.Graph()
B.add_nodes_from(df['Common name'].unique(), bipartite=0) # Add one set of nodes
B.add_nodes_from(df['Keyword'].unique(), bipartite=1) # Add the other set of nodes
B.add_edges_from([(row['Common name'], row['Keyword']) for idx, row in df.iterrows()]) # Add edges

# Get the two sets of nodes
top_nodes = {n for n, d in B.nodes(data=True) if d['bipartite']==0}
bottom_nodes = set(B) - top_nodes

# Compute the bipartite projection of the graph
G_top = nx.bipartite.projected_graph(B, top_nodes)
G_bottom = nx.bipartite.projected_graph(B, bottom_nodes)

# Calculate Common Neighbors (CN)
cn_top = [(e[0], e[1], len(list(nx.common_neighbors(G_top, e[0], e[1])))) for e in nx.non_edges(G_top)]
cn_bottom = [(e[0], e[1], len(list(nx.common_neighbors(G_bottom, e[0], e[1])))) for e in nx.non_edges(G_bottom)]

# Calculate Jaccard Coefficient (JC)
jc_top = list(nx.jaccard_coefficient(G_top))
jc_bottom = list(nx.jaccard_coefficient(G_bottom))

# Calculate Preferential AttachmentPA)
pa_top = list(nx.preferential_attachment(G_top))
pa_bottom = list(nx.preferential_attachment(G_bottom))

# Calculate Adamic-Adar (AA)
aa_top = list(nx.adamic_adar_index(G_top))
aa_bottom = list(nx.adamic_adar_index(G_bottom))

# Print the results
with open('Bipartiete Link Prediction Results.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow('Common Neighbors:')
        writer.writerow('\nTop:')
        writer.writerow(cn_top)
        writer.writerow('\nBottom')
        writer.writerow(cn_bottom)

        writer.writerow('\n\nJaccard Coefficient:')
        writer.writerow('\nTop:')
        writer.writerow(jc_top)
        writer.writerow('\nBottom')
        writer.writerow(jc_bottom)

        writer.writerow('\n\nPreferential Attachment:')
        writer.writerow('\nTop:')
        writer.writerow(pa_top)
        writer.writerow('\nBottom')
        writer.writerow(pa_bottom)

        writer.writerow('\n\nAdamic-Adar:')
        writer.writerow('\nTop:')
        writer.writerow(aa_top)
        writer.writerow('\nBottom')
        writer.writerow(aa_bottom)
