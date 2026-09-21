const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = async (env, options) => {
  const isDev = options.mode === "development";

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
        chunks: ["taskpane"]
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"]
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
