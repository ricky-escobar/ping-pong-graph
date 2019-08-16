from collections import Counter
from math import sqrt
from graphviz import Graph
from openpyxl import load_workbook

wb = load_workbook("SBP10 - Ping Pong Game History.xlsx")
game_data = wb["Game Data"]
winners = [cell.value for cell in game_data['C'][3:] if cell.value is not None]
losers = [cell.value for cell in game_data['F'][3:] if cell.value is not None]

pair_counter = Counter(tuple(sorted(pair)) for pair in zip(winners, losers))
game_counter = Counter(winners + losers)

players = {p for pair, count in pair_counter.items() if count > 3 and not any(" - " in p for p in pair) for p in pair}

g = Graph()
g.attr("node", shape="box")
for player in players:
    dim = str(sqrt(game_counter[player]) / 7)
    g.node(player, width=dim, height=dim)
for pair, count in pair_counter.items():
    if all(p in players for p in pair):
        g.edge(*pair, weight=str(count), penwidth=str(int(count ** .75)))

g.render("ping-pong", directory="out", format="svg")
