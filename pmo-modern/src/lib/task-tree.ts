// Utilitários para construir/iterar a árvore hierárquica de tarefas.

export type TaskNodeBase = {
  id: string;
  parentId: string | null;
  wbs: string | null;
  name: string;
  startDate: Date | string;
  endDate: Date | string;
  durationDays: number | null;
  percentDone: number;
  status: string;
  isMilestone: boolean;
  isSummary: boolean;
  deletedAt: Date | string | null;
};

export type TaskNode<T extends TaskNodeBase> = T & {
  children: TaskNode<T>[];
  depth: number;
};

export function buildTree<T extends TaskNodeBase>(rows: T[]): TaskNode<T>[] {
  const nodes = new Map<string, TaskNode<T>>();
  for (const r of rows) {
    nodes.set(r.id, { ...r, children: [], depth: 0 });
  }
  const roots: TaskNode<T>[] = [];
  for (const node of nodes.values()) {
    if (node.parentId && nodes.has(node.parentId)) {
      const parent = nodes.get(node.parentId)!;
      node.depth = parent.depth + 1;
      parent.children.push(node);
    } else {
      roots.push(node);
    }
  }
  // ordena recursivamente por wbs ou startDate
  const sorter = (a: TaskNode<T>, b: TaskNode<T>) => {
    if (a.wbs && b.wbs) return a.wbs.localeCompare(b.wbs, undefined, { numeric: true });
    return new Date(a.startDate).getTime() - new Date(b.startDate).getTime();
  };
  function sortRec(list: TaskNode<T>[]) {
    list.sort(sorter);
    list.forEach((n) => sortRec(n.children));
  }
  sortRec(roots);
  // recalcula depth pós-sort
  function setDepth(list: TaskNode<T>[], depth: number) {
    for (const n of list) {
      n.depth = depth;
      setDepth(n.children, depth + 1);
    }
  }
  setDepth(roots, 0);
  return roots;
}

export function flattenTree<T extends TaskNodeBase>(roots: TaskNode<T>[]): TaskNode<T>[] {
  const out: TaskNode<T>[] = [];
  function visit(list: TaskNode<T>[]) {
    for (const n of list) {
      out.push(n);
      visit(n.children);
    }
  }
  visit(roots);
  return out;
}
