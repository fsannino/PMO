import { PrismaClient } from "@prisma/client";
import bcrypt from "bcryptjs";

const prisma = new PrismaClient();

async function main() {
  const email = "fsannino@clbz.com";
  const newPassword = "admin123";
  const hash = await bcrypt.hash(newPassword, 10);

  const before = await prisma.user.findUnique({ where: { email } });
  if (!before) {
    console.log("ERRO: usuario nao existe");
    return;
  }
  console.log("Antes — active:", before.active, "role:", before.role);

  const updated = await prisma.user.update({
    where: { email },
    data: { passwordHash: hash, active: true },
  });
  console.log("Senha resetada. Hash novo gravado.");

  // Verifica que bate
  const ok = await bcrypt.compare(newPassword, updated.passwordHash);
  console.log("Senha 'admin123' bate?", ok);
}

main().finally(() => prisma.$disconnect());
