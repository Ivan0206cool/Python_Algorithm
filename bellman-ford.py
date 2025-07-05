def bellman_ford(graph, start_node):
    """
    Implements the Bellman-Ford algorithm to find shortest paths from a source node
    in a weighted directed graph.

    Args:
        graph (list of tuples): A list of tuples representing the edges in the graph.
                                Each tuple should be (u, v, weight), where u is the
                                source node, v is the destination node, and weight
                                is the weight of the edge.
        start_node: The starting node for finding shortest paths.

    Returns:
        dict: A dictionary where keys are nodes and values are their shortest
              distances from the start_node. Returns None if a negative-weight
              cycle is detected.
    """

    # Initialize distances:
    # Set distance to start_node as 0, and all other nodes as infinity.
    distances = {}
    for u, v, weight in graph:
        if u not in distances:
            distances[u] = float('inf')
        if v not in distances:
            distances[v] = float('inf')
    distances[start_node] = 0

    # Get the number of nodes in the graph
    # This is important for the relaxation steps.
    # We iterate N-1 times, where N is the number of nodes.
    num_nodes = len(distances)

    # Relax edges repeatedly N-1 times
    for _ in range(num_nodes - 1):
        for u, v, weight in graph:
            if distances[u] != float('inf') and distances[u] + weight < distances[v]:
                distances[v] = distances[u] + weight

    # Check for negative-weight cycles
    # After N-1 relaxations, if we can still relax any edge, it means
    # there's a negative-weight cycle.
    for u, v, weight in graph:
        if distances[u] != float('inf') and distances[u] + weight < distances[v]:
            print("Graph contains a negative-weight cycle!")
            return None  # Indicate that a negative cycle was found

    return distances

# Example Usage:

# Example 1: Graph with no negative cycles
print("--- Example 1: Graph with no negative cycles ---")
graph1 = [
    ('A', 'B', -1),
    ('A', 'C', 4),
    ('B', 'C', 3),
    ('B', 'D', 2),
    ('B', 'E', 2),
    ('D', 'B', 1),
    ('E', 'D', -3)
]
start_node1 = 'A'
shortest_paths1 = bellman_ford(graph1, start_node1)

if shortest_paths1:
    print(f"Shortest paths from {start_node1}:")
    for node, dist in shortest_paths1.items():
        print(f"  {node}: {dist}")

# Example 2: Graph with a negative cycle
print("\n--- Example 2: Graph with a negative cycle ---")
graph2 = [
    ('A', 'B', 1),
    ('B', 'C', -1),
    ('C', 'A', -1)  # This creates a negative cycle A -> B -> C -> A (1 - 1 - 1 = -1)
]
start_node2 = 'A'
shortest_paths2 = bellman_ford(graph2, start_node2)

if shortest_paths2:
    print(f"Shortest paths from {start_node2}:")
    for node, dist in shortest_paths2.items():
        print(f"  {node}: {dist}")

# Example 3: Another graph
print("\n--- Example 3: Another graph ---")
graph3 = [
    ('S', 'E', 8),
    ('S', 'A', 10),
    ('E', 'D', 1),
    ('D', 'A', -4),
    ('A', 'C', 2),
    ('C', 'B', -1),
    ('B', 'A', 1),
    ('D', 'C', -1)
]
start_node3 = 'S'
shortest_paths3 = bellman_ford(graph3, start_node3)

if shortest_paths3:
    print(f"Shortest paths from {start_node3}:")
    for node, dist in shortest_paths3.items():
        print(f"  {node}: {dist}")