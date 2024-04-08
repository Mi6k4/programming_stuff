from python.veryday_routine.all_dep_list import edges

from igraph import Graph

# Create an undirected graph
g = Graph(directed=True)


# Create vertices and add edges
vertices = list(set([v for e in edges for v in e]))
g.add_vertices(vertices)
g.add_edges(edges)

# Check for cycles
if g.is_cyclic():
    print("The graph contains a cycle.")
else:
    print("The graph is acyclic.")






