import { prisma } from "./db";

/**
 * Verifica se adicionar a aresta predecessor → successor criaria um ciclo
 * no grafo de dependências do projeto.
 * Retorna true se há ciclo (não pode criar a aresta).
 */
export async function wouldCreateCycle(
  predecessorId: string,
  successorId: string,
): Promise<boolean> {
  if (predecessorId === successorId) return true;

  // BFS partindo do successor seguindo as arestas existentes; se chegarmos
  // ao predecessor, há ciclo.
  const visited = new Set<string>([successorId]);
  const queue: string[] = [successorId];

  while (queue.length) {
    const current = queue.shift()!;
    const next = await prisma.taskDependency.findMany({
      where: { predecessorId: current },
      select: { successorId: true },
    });
    for (const { successorId: sid } of next) {
      if (sid === predecessorId) return true;
      if (!visited.has(sid)) {
        visited.add(sid);
        queue.push(sid);
      }
    }
  }
  return false;
}
