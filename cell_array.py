class CellArray:
    def __init__(self, x0, y0, width, height):
        assert x0 > 0 and y0 > 0 and width > 0 and height > 0, "x0, y0, width and height must be 1 or greater"
        self.x0 = x0
        self.y0 = y0
        self.width = width
        self.height = height
        self.matrix = self._populate_matrix(x0, y0, width, height)

    def __iter__(self):
        coordinates = [coordinate for row in self.matrix for coordinate in row]
        return iter(coordinates)

    def get_all_edges(self, top=True, right=True, bottom=True, left=True):
        edges = []
        for cell in self:
            cell_borders = self._get_cell_borders(cell, top, right, bottom, left)

            if self._is_on_edge(cell_borders):
                edges.append({
                    "position": cell,
                    "borders": cell_borders
                })

        return edges

    def _get_cell_borders(self, cell, top=True, right=True, bottom=True, left=True):
        return (
            1 if left and self._is_left_edge(cell) else 0,
            1 if right and self._is_right_edge(cell) else 0,
            1 if top and self._is_top_edge(cell) else 0,
            1 if bottom and self._is_bottom_edge(cell) else 0
        )

    def _is_top_edge(self, cell):
        return True if cell[1] == self.y0 else False

    def _is_bottom_edge(self, cell):
        return True if cell[1] == self.y0 + self.height - 1 else False

    def _is_left_edge(self, cell):
        return True if cell[0] == self.x0 else False

    def _is_right_edge(self, cell):
        return True if cell[0] == self.x0 + self.width - 1 else False

    @staticmethod
    def _is_on_edge(borders):
        return True if 1 in borders else False

    @staticmethod
    def _populate_matrix(x0, y0, width, height):
        return [[(i, j) for i in range(x0, x0 + width)] for j in range(y0, y0 + height)]
