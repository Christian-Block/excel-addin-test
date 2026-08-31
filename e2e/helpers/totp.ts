import { createHmac } from "crypto";

/**
 * Generates a time based one time password (RFC 6238) for the multi factor
 * authentication of the Microsoft sign in flow.
 *
 * The test account has to use the "authenticator app" method; its shared
 * secret is provided through the `E2E_M365_TOTP_SECRET` environment variable.
 */
export function generateTotp(base32Secret: string, forTime: number = Date.now()): string {
  const key = base32Decode(base32Secret);
  const counter = Math.floor(forTime / 1000 / 30);
  const counterBuffer = Buffer.alloc(8);

  counterBuffer.writeUInt32BE(Math.floor(counter / 0x100000000), 0);
  counterBuffer.writeUInt32BE(counter >>> 0, 4);

  const digest = createHmac("sha1", key).update(counterBuffer).digest();
  const offset = digest[digest.length - 1] & 0x0f;
  const binary =
    ((digest[offset] & 0x7f) << 24) | (digest[offset + 1] << 16) | (digest[offset + 2] << 8) | digest[offset + 3];

  return String(binary % 1_000_000).padStart(6, "0");
}

function base32Decode(input: string): Buffer {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
  const cleaned = input.replace(/=+$/, "").replace(/\s/g, "").toUpperCase();
  const bytes: number[] = [];
  let bits = 0;
  let value = 0;

  for (const character of cleaned) {
    const index = alphabet.indexOf(character);

    if (index === -1) {
      throw new Error(`'${character}' is not a valid base32 character.`);
    }

    value = (value << 5) | index;
    bits += 5;

    if (bits >= 8) {
      bits -= 8;
      bytes.push((value >>> bits) & 0xff);
    }
  }

  return Buffer.from(bytes);
}
