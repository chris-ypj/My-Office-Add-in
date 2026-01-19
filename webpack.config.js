/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  return {
    devtool: "source-map",

    // ✅ Only JS entries here (don't include HTML in entry)
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      // dialog page can be pure HTML; no JS bundle needed
      // If you later add src/dialog/dialog.js, then add: dialog: "./src/dialog/dialog.js",
    },

    output: {
      clean: true,
    },

    resolve: {
      extensions: [".html", ".js"],
    },

    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: { loader: "babel-loader" },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },

    plugins: [
      // ✅ Output to /taskpane/taskpane.html
      new HtmlWebpackPlugin({
        filename: "taskpane/taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),

      // ✅ Output to /commands/commands.html
      new HtmlWebpackPlugin({
        filename: "commands/commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),

      // ✅ Output to /dialog/dialog.html (no JS bundle required)
      new HtmlWebpackPlugin({
        filename: "dialog/dialog.html",
        template: "./src/dialog/dialog.html",
        chunks: [],
      }),

      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              if (dev) return content;
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
    ],

    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      // Office web clients often block localhost WS; disable HMR/WS to avoid errors.
      webSocketServer: false,
      server: {
        type: "https",
        options:
            env.WEBPACK_BUILD || options.https !== undefined
                ? options.https
                : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };
};
