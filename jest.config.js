/* eslint-env node */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "jsdom",
  testEnvironmentOptions: { customExportConditions: ["node"] },
  roots: ["<rootDir>/test"],
  testMatch: ["**/*.test.ts"],
  collectCoverageFrom: ["src/**/*.ts"],
  transform: {
    "^.+\\.ts$": ["ts-jest", { tsconfig: { module: "commonjs", types: ["office-js", "custom-functions-runtime", "jest", "node"] } }],
  },
};
