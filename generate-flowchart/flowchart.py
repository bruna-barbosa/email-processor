import json
from graphviz import Digraph
def create_flowchart(flowchart_data):
    
# Define pastel colors
    pastel_colors = ['#FAD6A5', '#F6F6D8', '#CBF1F5', '#A3D9E8', '#E8B4BC', '#B0E57C', '#D0C7A8']
    
# Define some global graph attributes
    graph = Digraph('G', filename='flowchart', format='png', graph_attr={'rankdir': 'LR', 'splines': 'ortho', 'bgcolor': 'white'})
    
# Add nodes (steps) to the flowchart with rectangular shape
    for idx, step in enumerate(flowchart_data["steps"]):
        node_color = pastel_colors[idx % len(pastel_colors)]
        if step["id"] == 1:  # Start node
            graph.node(str(step["id"]), step["label"], shape='box', style='filled', fillcolor=node_color, fontname='Consolas', fontsize='14')
        elif step["id"] == len(flowchart_data["steps"]):  # End node
            graph.node(str(step["id"]), step["label"], shape='box', style='filled', fillcolor=node_color, fontname='Consolas', fontsize='14')
        else:
            graph.node(str(step["id"]), step["label"], shape='box', style='filled', fillcolor=node_color, fontname='Consolas', fontsize='12')
    
# Add connections between nodes with custom labels and straight arrows
    for connection in flowchart_data["connections"]:
        graph.edge(str(connection["from"]), str(connection["to"]), xlabel=connection.get("label", ""), fontsize='10', color='black', fontname='Consolas', dir='forward', arrowhead='normal', arrowsize='0.5', labelloc='t', fontcolor='black', fontbgcolor='white', xlabeldistance='2')
   
# Add error handling paths with specific styling and labels
    for idx, error_path in enumerate(flowchart_data["error_handling"]):
        graph.edge(str(error_path["from"]), str(error_path["to"]), xlabel=error_path["label"], fillcolor='white', fontsize='10', color='red', fontname='Consolas', style="dashed",  arrowsize='0.5', labelloc="t", fontcolor='red', fontbgcolor='white', xlabeldistance='5', penwidth='1')
        if idx > 0:
            graph.attr("edge", nodesep='3')  # Add more space to the vertical lines

    
# Comment nodes for additional explanations
    graph.node("comment1", label="Flowchart Automation by bruna.duarte@nokia.com. Version 1: July 2023. ", shape="note", fontsize='10', fontname='Consolas', color='gray')
    graph.edge("comment1", str(flowchart_data["steps"][1]["id"]), style="invis")  # Example: Connect comment to a specific node
    return graph

if __name__ == "__main__":
    # Read flowchart outline data from JSON file
    with open("flowchart_outline.json", "r") as f:
        flowchart_data = json.load(f)
    # Generate the flowchart and save it as a PNG image
    flowchart = create_flowchart(flowchart_data)
    # Apply some beautification options
    flowchart.attr(dpi='420')
    # Set global node attributes
    flowchart.attr('node', fontname='Consolas', style='filled', fillcolor='white', fontcolor='black', fixedsize='true', width='1.7', height='1.7')
    # Set global edge attributes
    flowchart.attr('edge', fontname='Consolas')
    flowchart.render(view=True)
