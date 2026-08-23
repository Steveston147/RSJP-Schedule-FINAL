import { randomBytes, scryptSync } from "node:crypto";
import { createInterface } from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";

const rl = createInterface({ input, output });
const password = await rl.question("PoC password: ");
rl.close();

if (!password) {
  console.error("Password must not be empty.");
  process.exit(1);
}

const salt = randomBytes(16).toString("hex");
const hash = scryptSync(password, salt, 64).toString("hex");
console.log(`scrypt$${salt}$${hash}`);
