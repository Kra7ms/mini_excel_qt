from collections import defaultdict, deque
from typing import Set, Dict, Iterable
from dependency import DependencyExtractor


class CircularDependencyError(Exception):
    pass


class DependencyGraph:
    """
    Hücreler arası bağımlılık grafiği
    A1 -> {B1, C1}
    """
    def __init__(self):
        # cell -> set of cells it depends on
        self.dependencies: Dict[str, Set[str]] = defaultdict(set)

        # reverse graph: cell -> cells depending on it
        self.dependents: Dict[str, Set[str]] = defaultdict(set)

        # A -> {B, C}  (A değişirse B ve C etkilenir)
        self.forward = defaultdict(set)

        # B -> {A}     (B, A'ya bağlı)
        self.reverse = defaultdict(set)

        self.extractor = DependencyExtractor()

    # =====================================================
    # DEPENDENCY EKLEME
    # =====================================================
    def set_dependencies(self, cell: str, deps: Iterable[str]):
        """
        cell: "A1"
        deps: {"B1", "C1"}
        """
        self.remove_cell(cell)

        for dep in deps:
            self.dependencies[cell].add(dep)
            self.dependents[dep].add(cell)

    def remove_cell(self, cell: str):
        """
        Hücreyi grafikten tamamen çıkar
        """
        if cell in self.dependencies:
            for dep in self.dependencies[cell]:
                self.dependents[dep].discard(cell)
            del self.dependencies[cell]

        if cell in self.dependents:
            del self.dependents[cell]

    # =====================================================
    # RE-CALCULATE
    # =====================================================
    def recalculate_dependents(self, start_cell, engine):
        """
        start_cell değişti → etkilenen tüm hücreleri yeniden hesapla
        """
        visited = set()
        queue = deque([start_cell])

        while queue:
            current = queue.popleft()

            for dependent in self.forward[current]:
                if dependent in visited:
                    continue

                visited.add(dependent)

                r, c = dependent
                engine.recalculate_cell(r, c)

                queue.append(dependent)

    # =====================================================
    # SORGULAR
    # =====================================================
    def get_dependencies(self, cell: str) -> Set[str]:
        return self.dependencies.get(cell, set())

    def get_dependents(self, cell: str) -> Set[str]:
        return self.dependents.get(cell, set())

    # =====================================================
    # TOPOLOGICAL SORT
    # =====================================================
    def evaluation_order(self, start_cells: Iterable[str] = None) -> list:
        """
        Yeniden hesaplama sırası üretir
        """
        indegree = defaultdict(int)

        for cell, deps in self.dependencies.items():
            indegree.setdefault(cell, 0)
            for d in deps:
                indegree[cell] += 1
                indegree.setdefault(d, 0)

        queue = deque(c for c, deg in indegree.items() if deg == 0)
        order = []

        while queue:
            node = queue.popleft()
            order.append(node)

            for dependent in self.dependents.get(node, []):
                indegree[dependent] -= 1
                if indegree[dependent] == 0:
                    queue.append(dependent)

        if len(order) != len(indegree):
            raise CircularDependencyError("Circular dependency detected")

        if start_cells:
            return [c for c in order if c in start_cells or self._depends_on(c, start_cells)]

        return order

    # =====================================================
    # YARDIMCI
    # =====================================================
    def _depends_on(self, cell: str, targets: Iterable[str]) -> bool:
        visited = set()
        stack = [cell]

        while stack:
            current = stack.pop()
            if current in visited:
                continue
            visited.add(current)

            for dep in self.dependencies.get(current, []):
                if dep in targets:
                    return True
                stack.append(dep)

        return False

    # =====================================================
    # DEBUG
    # =====================================================
    def dump(self):
        return {
            "dependencies": dict(self.dependencies),
            "dependents": dict(self.dependents)
        }
