/* eslint-env node */
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");

const urlProd = "https://christian-block.github.io/excel-addin-test/";
const urlDev = "https://localhost:3000/";

module.exports = (env, options) => {
  const dev = options.mode !== "production";

  return {
    devtool: "source-map",
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts",
      functions: "./src/functions/functions.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".js"],
    },
    module: {
      rules: [
        { test: /\.ts$/, exclude: /node_modules/, use: "ts-loader" },
        { test: /\.(png|jpg|gif|ico)$/, type: "asset/resource", generator: { filename: "assets/[name][ext]" } },
      ],
    },
    plugins: [
      new CustomFunctionsMetadataPlugin({ output: "functions.json", input: "./src/functions/functions.ts" }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform(content) {
              return dev ? content.toString().replace(new RegExp(urlProd, "g"), urlDev) : content;
            },
          },
        ],
      }),
    ],
    devServer: {
      static: { directory: path.join(__dirname, "dist") },
      server: "https",
      port: 3000,
      headers: { "Access-Control-Allow-Origin": "*" },
    },
  };
};
