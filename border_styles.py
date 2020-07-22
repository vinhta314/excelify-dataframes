from itertools import product
from openpyxl.styles import Border, Side


class BorderStyle:
    def __init__(self, thickness):
        self.style_map = self._set_style_map(thickness)

    def __getitem__(self, item):
        return self.style_map[item]

    def _set_style_map(self, thickness):
        return {edges: self._create_border(edges, thickness) for edges in self._get_edge_permutations()}

    @staticmethod
    def _create_border(edges, thickness):
        side_styles = {}

        if edges[0] == 1:
            side_styles['left'] = Side(style=thickness)
        if edges[1] == 1:
            side_styles['right'] = Side(style=thickness)
        if edges[2] == 1:
            side_styles['top'] = Side(style=thickness)
        if edges[3] == 1:
            side_styles['bottom'] = Side(style=thickness)

        return Border(**side_styles)

    @staticmethod
    def _get_edge_permutations():
        return [perm for perm in product((0, 1), repeat=4)][1:]
