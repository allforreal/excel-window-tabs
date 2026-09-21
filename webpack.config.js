const path = require("path");
const fs = require("fs");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

function readManifestVersion() {
  try {
    const manifest = fs.readFileSync(path.resolve(__dirname, "manifest.xml"), "utf8");
    return manifest.match(/<Version>\s*([^<]+?)\s*<\/Version>/i)?.[1] ?? "0";
  } catch {
    return "0";
  }
}

module.exports = async (env, options) => {
  const isDev = options.mode === "development";
  const version = readManifestVersion();

  const config = {
    devtool: isDev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts"
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
      publicPath: ""
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"]
    },
    module: {
      rules: [{ test: /\.tsx?$/, exclude: /node_modules/, use: "ts-loader" }]
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        inject: false,
        templateParameters: { version }
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
        inject: false,
        templateParameters: { version }
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" }
        ]
      })
    ]
  };

  if (isDev) {
    const devCerts = require("office-addin-dev-certs");
    const httpsOptions = await devCerts.getHttpsServerOptions();
    config.devServer = {
      static: { directory: path.join(__dirname, "dist") },
      headers: { "Access-Control-Allow-Origin": "*" },
      server: {
        type: "https",
        options: {
          key: httpsOptions.key,
          cert: httpsOptions.cert,
          ca: httpsOptions.ca
        }
      },
      port: 3000,
      hot: true
    };
  }

  return config;
};
