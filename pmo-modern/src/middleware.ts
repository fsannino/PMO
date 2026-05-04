// Protege todas as rotas autenticadas redirecionando para /login quando não há JWT.
// As verificações de RBAC mais finas (Access matrix por projeto/módulo) são feitas
// nos próprios layouts/route handlers via src/lib/access.ts.

import { withAuth } from "next-auth/middleware";

export default withAuth({
  pages: { signIn: "/login" },
});

export const config = {
  matcher: [
    // Tudo, exceto rotas públicas e estáticos:
    "/((?!login|api/auth|_next/static|_next/image|favicon.ico|public).*)",
  ],
};
