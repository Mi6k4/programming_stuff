from collections import defaultdict
from python.veryday_routine.all_dep_list import edges

class Graph:
    def __init__(self):
        self.graph = defaultdict(list)

    def add_edge(self, u, v):
        self.graph[u].append(v)

    def is_cyclic_util(self, v, visited, stack):
        visited[v] = True
        stack[v] = True

        for neighbor in self.graph[v]:
            if not visited[neighbor]:
                if self.is_cyclic_util(neighbor, visited, stack):
                    return True
            elif stack[neighbor]:
                return True

        stack[v] = False
        return False

    def is_cyclic(self):
        visited = {node: False for node in self.graph}
        stack = {node: False for node in self.graph}

        for node in self.graph:
            if not visited[node]:
                if self.is_cyclic_util(node, visited, stack):
                    return True

        return False
# Create the graph
g = Graph()


for edge in edges:
    g.add_edge(edge[0], edge[1])

if g.is_cyclic():
    print("The graph contains a cycle.")
else:
    print("The graph is acyclic.")