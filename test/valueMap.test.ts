import {
  DEFAULT_VALUE_MAP,
  getValueMap,
  loadValueMapFromEndpoint,
  lookupValue,
  parseValueMap,
  resetValueMap,
  setValueMap,
} from "../src/functions/valueMap";
import { loadMap, mapSize, replaceRange, replaceValue } from "../src/functions/functions";

describe("value map", () => {
  beforeEach(() => resetValueMap());

  it("uses the list that ships with the add-in", () => {
    expect(getValueMap()).toEqual(DEFAULT_VALUE_MAP);
    expect(lookupValue("DE")).toBe("Germany");
  });

  it("looks keys up case insensitively and ignores whitespace", () => {
    expect(lookupValue(" de ")).toBe("Germany");
  });

  it("returns the key or the fallback for unknown keys", () => {
    expect(lookupValue("XX")).toBe("XX");
    expect(lookupValue("XX", "unknown")).toBe("unknown");
    expect(lookupValue(null)).toBe("");
  });

  it("can be replaced at runtime", () => {
    setValueMap({ A: "Alpha" });

    expect(lookupValue("A")).toBe("Alpha");
    expect(lookupValue("DE")).toBe("DE");
  });

  it("parses both supported REST payload shapes", () => {
    expect(parseValueMap({ DE: "Germany" })).toEqual({ DE: "Germany" });
    expect(parseValueMap([{ key: "DE", value: "Germany" }])).toEqual({ DE: "Germany" });
  });

  it("rejects invalid REST payloads", () => {
    expect(() => parseValueMap("nope")).toThrow(/object or an array/);
    expect(() => parseValueMap({ DE: 1 })).toThrow(/must be a string/);
    expect(() => parseValueMap([{ key: "DE" }])).toThrow(/'key' and 'value'/);
  });

  it("loads the map from a REST endpoint", async () => {
    const fetchMock = jest.fn().mockResolvedValue({ ok: true, json: async () => ({ A: "Alpha", B: "Beta" }) });

    await expect(loadValueMapFromEndpoint("https://example.com/map.json", fetchMock as never)).resolves.toBe(2);
    expect(fetchMock).toHaveBeenCalledWith("https://example.com/map.json");
    expect(lookupValue("B")).toBe("Beta");
  });

  it("reports HTTP errors of the endpoint", async () => {
    const fetchMock = jest.fn().mockResolvedValue({ ok: false, status: 404, json: async () => ({}) });

    await expect(loadValueMapFromEndpoint("https://example.com/map.json", fetchMock as never)).rejects.toThrow(
      /HTTP 404/
    );
  });
});

describe("custom functions", () => {
  beforeEach(() => resetValueMap());

  it("replaces a single value", () => {
    expect(replaceValue("DE")).toBe("Germany");
    expect(replaceValue("XX", "n/a")).toBe("n/a");
  });

  it("replaces every value of a range", () => {
    expect(replaceRange([["DE", "FR"], ["XX", "US"]], "n/a")).toEqual([
      ["Germany", "France"],
      ["n/a", "United States"],
    ]);
  });

  it("reports the size of the current map", () => {
    expect(mapSize()).toBe(Object.keys(DEFAULT_VALUE_MAP).length);
  });

  it("loads the map from an endpoint", async () => {
    const fetchMock = jest.fn().mockResolvedValue({ ok: true, json: async () => ({ A: "Alpha" }) });
    (globalThis as unknown as { fetch: unknown }).fetch = fetchMock;

    await expect(loadMap("https://example.com/map.json")).resolves.toBe(1);
    expect(mapSize()).toBe(1);
  });
});
