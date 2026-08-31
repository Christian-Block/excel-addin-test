/**
 * Very small authentication layer used to separate the free feature area from
 * the features that are only available to signed in users.
 *
 * The add-in uses single sign-on (`Office.auth.getAccessToken`). The token is
 * kept in memory only; it is never written to the document or to storage.
 */

export type FeatureId =
  | "create-table"
  | "read-write-cells"
  | "validate-values"
  | "row-conditional-format"
  | "column-conditional-format"
  | "json-export"
  | "json-import"
  | "user-settings";

export interface FeatureDefinition {
  id: FeatureId;
  title: string;
  requiresAuthentication: boolean;
}

export const FEATURES: FeatureDefinition[] = [
  { id: "create-table", title: "Create demo table", requiresAuthentication: false },
  { id: "read-write-cells", title: "Read and write cells", requiresAuthentication: false },
  { id: "validate-values", title: "Validate table values", requiresAuthentication: false },
  { id: "row-conditional-format", title: "Highlight rows", requiresAuthentication: false },
  { id: "column-conditional-format", title: "Colour scale for a column", requiresAuthentication: true },
  { id: "json-export", title: "Download data as JSON", requiresAuthentication: false },
  { id: "json-import", title: "Import data from JSON", requiresAuthentication: true },
  { id: "user-settings", title: "Store user settings", requiresAuthentication: true },
];

export function getFeature(id: FeatureId): FeatureDefinition {
  const feature = FEATURES.find((candidate) => candidate.id === id);

  if (!feature) {
    throw new Error(`Unknown feature '${id}'.`);
  }

  return feature;
}

export class AuthService {
  private accessToken: string | undefined;

  public get isSignedIn(): boolean {
    return this.accessToken !== undefined;
  }

  /**
   * Requests an access token from the Office host and marks the user as
   * signed in.
   */
  public async signIn(): Promise<void> {
    const token = await Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true });

    if (!token) {
      throw new Error("The Office host did not return an access token.");
    }

    this.accessToken = token;
  }

  public signOut(): void {
    this.accessToken = undefined;
  }

  /**
   * Returns the current access token, for example to call a protected API.
   */
  public getAccessToken(): string | undefined {
    return this.accessToken;
  }

  public isFeatureAvailable(id: FeatureId): boolean {
    return !getFeature(id).requiresAuthentication || this.isSignedIn;
  }

  /**
   * Throws when the feature may not be used with the current sign in state.
   */
  public assertFeatureAvailable(id: FeatureId): void {
    if (!this.isFeatureAvailable(id)) {
      throw new Error(`'${getFeature(id).title}' is only available for signed in users.`);
    }
  }
}
