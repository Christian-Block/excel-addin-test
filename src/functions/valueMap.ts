/**
 * Key/value map used by the custom functions to replace values.
 *
 * The map ships with the add-in today. `loadValueMapFromEndpoint` already
 * implements the planned REST based loading, so the built in list can be
 * replaced at runtime without touching the custom functions.
 */

export type ValueMap = Record<string, string>;

export const DEFAULT_VALUE_MAP: ValueMap = {
  DE: "Germany",
  AT: "Austria",
  CH: "Switzerland",
  FR: "France",
  IT: "Italy",
  US: "United States",
  GB: "United Kingdom",
};

let currentMap: ValueMap = { ...DEFAULT_VALUE_MAP };

export function getValueMap(): ValueMap {
  return { ...currentMap };
}

export function setValueMap(map: ValueMap): void {
  currentMap = { ...map };
}

export function resetValueMap(): void {
  currentMap = { ...DEFAULT_VALUE_MAP };
}

/**
 * Looks a key up in the current map. The lookup is case insensitive and
 * ignores surrounding whitespace.
 *
 * @param key Key to look up.
 * @param fallback Value returned when the key is unknown.
 */
export function lookupValue(key: unknown, fallback?: string): string {
  if (typeof key !== "string" && typeof key !== "number") {
    return fallback ?? "";
  }

  const normalized = String(key).trim().toLowerCase();
  const match = Object.keys(currentMap).find((candidate) => candidate.trim().toLowerCase() === normalized);

  if (match === undefined) {
    return fallback ?? String(key);
  }

  return currentMap[match];
}

/**
 * Validates a REST response body and converts it into a {@link ValueMap}.
 *
 * Two shapes are accepted:
 * - an object with string values, for example `{ "DE": "Germany" }`
 * - an array of `{ key, value }` entries
 *
 * @throws Error when the payload cannot be converted.
 */
export function parseValueMap(payload: unknown): ValueMap {
  if (Array.isArray(payload)) {
    const map: ValueMap = {};

    for (const entry of payload) {
      if (typeof entry !== "object" || entry === null) {
        throw new Error("Every entry of the value map must be an object with 'key' and 'value'.");
      }

      const { key, value } = entry as { key?: unknown; value?: unknown };

      if (typeof key !== "string" || typeof value !== "string") {
        throw new Error("Every entry of the value map must be an object with 'key' and 'value'.");
      }

      map[key] = value;
    }

    return map;
  }

  if (typeof payload === "object" && payload !== null) {
    const map: ValueMap = {};

    for (const [key, value] of Object.entries(payload)) {
      if (typeof value !== "string") {
        throw new Error(`The value for key '${key}' must be a string.`);
      }

      map[key] = value;
    }

    return map;
  }

  throw new Error("The value map endpoint must return an object or an array.");
}

/**
 * Loads the key/value map from a REST endpoint and makes it the current map.
 *
 * @param url Endpoint returning the map.
 * @param fetchImpl Injectable fetch implementation, used by the unit tests.
 * @returns The number of entries that were loaded.
 */
export async function loadValueMapFromEndpoint(
  url: string,
  fetchImpl: typeof fetch = fetch
): Promise<number> {
  const response = await fetchImpl(url);

  if (!response.ok) {
    throw new Error(`The value map could not be loaded (HTTP ${response.status}).`);
  }

  const map = parseValueMap(await response.json());

  setValueMap(map);

  return Object.keys(map).length;
}
