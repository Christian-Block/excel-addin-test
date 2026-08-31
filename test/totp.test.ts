import { generateTotp } from "../e2e/helpers/totp";

// Test vector of RFC 6238 (SHA-1, secret "12345678901234567890").
const SECRET = "GEZDGNBVGY3TQOJQGEZDGNBVGY3TQOJQ";

describe("generateTotp", () => {
  it("creates the codes of the RFC 6238 test vector", () => {
    expect(generateTotp(SECRET, 59_000)).toBe("287082");
    expect(generateTotp(SECRET, 1_111_111_109_000)).toBe("081804");
    expect(generateTotp(SECRET, 1_234_567_890_000)).toBe("005924");
  });

  it("creates codes with six digits", () => {
    expect(generateTotp(SECRET)).toMatch(/^\d{6}$/);
  });

  it("rejects invalid base32 secrets", () => {
    expect(() => generateTotp("not-base32!")).toThrow(/base32/);
  });
});
