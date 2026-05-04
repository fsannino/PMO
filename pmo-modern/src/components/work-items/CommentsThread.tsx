"use client";

import { useState, useTransition } from "react";
import { Trash2, Send } from "lucide-react";
import { Button } from "@/components/ui/Button";
import { Textarea } from "@/components/ui/Field";
import { addComment, deleteComment } from "@/lib/actions/work-items";

export type CommentItem = {
  id: string;
  body: string;
  createdAt: Date | string;
  authorId: string;
  author: { name: string };
};

type Target =
  | { taskId: string }
  | { issueId: string }
  | { riskId: string }
  | { actionId: string }
  | { changeRequestId: string };

export function CommentsThread({
  comments,
  target,
  currentUserId,
  isAdmin,
}: {
  comments: CommentItem[];
  target: Target;
  currentUserId: string;
  isAdmin: boolean;
}) {
  const [body, setBody] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [pending, start] = useTransition();

  function submit() {
    if (!body.trim()) return;
    setError(null);
    start(async () => {
      const res = await addComment({ body, ...target });
      if (res.ok) setBody("");
      else setError(res.error);
    });
  }
  function remove(id: string) {
    if (!confirm("Apagar comentário?")) return;
    setError(null);
    start(async () => {
      const res = await deleteComment(id);
      if (!res.ok) setError(res.error);
    });
  }

  return (
    <div className="space-y-3">
      <h4 className="text-sm font-semibold uppercase tracking-wide text-slate-500">
        Comentários ({comments.length})
      </h4>

      {error && (
        <div className="rounded border border-rose-200 bg-rose-50 px-3 py-2 text-xs text-rose-700">{error}</div>
      )}

      <ul className="space-y-2">
        {comments.length === 0 && (
          <li className="rounded border border-dashed border-slate-300 p-3 text-center text-xs text-slate-500">
            Sem comentários ainda.
          </li>
        )}
        {comments.map((c) => {
          const canDelete = c.authorId === currentUserId || isAdmin;
          return (
            <li key={c.id} className="rounded border bg-white p-3 text-sm">
              <div className="flex items-center justify-between gap-2">
                <span className="font-medium text-slate-900">{c.author.name}</span>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-slate-400">
                    {new Date(c.createdAt).toLocaleString("pt-BR")}
                  </span>
                  {canDelete && (
                    <button onClick={() => remove(c.id)} className="text-slate-300 hover:text-rose-600" title="Apagar">
                      <Trash2 className="h-3.5 w-3.5" />
                    </button>
                  )}
                </div>
              </div>
              <p className="mt-1 whitespace-pre-wrap text-slate-700">{c.body}</p>
            </li>
          );
        })}
      </ul>

      <div className="rounded border bg-slate-50 p-3">
        <Textarea
          rows={2}
          value={body}
          onChange={(e) => setBody(e.target.value)}
          placeholder="Escreva um comentário…"
          maxLength={4000}
        />
        <div className="mt-2 flex justify-end">
          <Button size="sm" onClick={submit} disabled={pending || !body.trim()}>
            <Send className="h-3.5 w-3.5" /> {pending ? "Enviando…" : "Comentar"}
          </Button>
        </div>
      </div>
    </div>
  );
}
