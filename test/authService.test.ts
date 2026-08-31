import { installOfficeMock } from "./mocks/officeMock";
import { AuthService, FEATURES, getFeature } from "../src/auth/authService";

describe("AuthService", () => {
  it("only allows free features while signed out", () => {
    const auth = new AuthService();

    expect(auth.isSignedIn).toBe(false);
    expect(auth.isFeatureAvailable("create-table")).toBe(true);
    expect(auth.isFeatureAvailable("json-import")).toBe(false);
    expect(() => auth.assertFeatureAvailable("user-settings")).toThrow(/signed in/);
  });

  it("unlocks the protected features after signing in", async () => {
    installOfficeMock();
    const auth = new AuthService();

    await auth.signIn();

    expect(auth.isSignedIn).toBe(true);
    expect(auth.getAccessToken()).toBe("access-token");
    for (const feature of FEATURES) {
      expect(auth.isFeatureAvailable(feature.id)).toBe(true);
    }
  });

  it("locks the protected features again after signing out", async () => {
    installOfficeMock();
    const auth = new AuthService();

    await auth.signIn();
    auth.signOut();

    expect(auth.getAccessToken()).toBeUndefined();
    expect(auth.isFeatureAvailable("column-conditional-format")).toBe(false);
  });

  it("fails when the host does not return a token", async () => {
    const officeMock = installOfficeMock();
    officeMock.auth.token = "";

    await expect(new AuthService().signIn()).rejects.toThrow(/access token/);
  });

  it("throws for unknown features", () => {
    expect(() => getFeature("nope" as never)).toThrow(/Unknown feature/);
  });
});
