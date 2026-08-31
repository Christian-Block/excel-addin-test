import { getValueMap, loadValueMapFromEndpoint, lookupValue } from "./valueMap";

/* global CustomFunctions */

/**
 * Replaces a key with the matching value from the key/value map of the add-in.
 * @customfunction REPLACE
 * @param key Key to replace, for example a country code.
 * @param [fallback] Value returned when the key is unknown.
 * @returns The mapped value.
 */
export function replaceValue(key: string, fallback?: string): string {
  return lookupValue(key, fallback);
}

/**
 * Replaces every key of a range with the matching value from the key/value map.
 * @customfunction REPLACERANGE
 * @param keys Range of keys to replace.
 * @param [fallback] Value returned for unknown keys.
 * @returns A matrix with the mapped values.
 */
export function replaceRange(keys: string[][], fallback?: string): string[][] {
  return keys.map((row) => row.map((key) => lookupValue(key, fallback)));
}

/**
 * Returns the number of entries in the current key/value map.
 * @customfunction MAPSIZE
 * @returns Number of entries.
 */
export function mapSize(): number {
  return Object.keys(getValueMap()).length;
}

/**
 * Loads the key/value map from a REST endpoint and returns the number of
 * entries that were loaded.
 * @customfunction LOADMAP
 * @param url Endpoint returning the key/value map.
 * @returns Number of entries that were loaded.
 */
export async function loadMap(url: string): Promise<number> {
  return loadValueMapFromEndpoint(url);
}

if (typeof CustomFunctions !== "undefined") {
  CustomFunctions.associate("REPLACE", replaceValue);
  CustomFunctions.associate("REPLACERANGE", replaceRange);
  CustomFunctions.associate("MAPSIZE", mapSize);
  CustomFunctions.associate("LOADMAP", loadMap);
}
