// Build for Nymform for Excel. Mirrors the Office add-in React task pane scaffold
// (webpack + ts-loader + HtmlWebpackPlugin + dev certs), with build-time config (spec §3)
// and the Content Security Policy (spec §7.13).
const path = require("path");
const { execSync } = require("child_process");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const pkg = require("./package.json");

// The dev server's port, set once in package.json ("config.dev_server_port"); the dev and bench
// manifests use the same port (tests/manifests.test.ts checks).
const DEV_PORT = pkg.config.dev_server_port;
const DEV_URL = `https://localhost:${DEV_PORT}/`;

function commitHash() {
  if (process.env.GITHUB_SHA) return process.env.GITHUB_SHA.slice(0, 7);
  try {
    return execSync("git rev-parse --short HEAD", { stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
  } catch {
    return "unknown";
  }
}

function csp(endpointOrigin) {
  return [
    "default-src 'self'",
    "script-src 'self' https://appsforoffice.microsoft.com",
    `connect-src ${endpointOrigin}`,
    "style-src 'self' 'unsafe-inline'",
    "img-src 'self' data:",
    "font-src 'self'",
    "object-src 'none'",
    "base-uri 'self'",
    // form-action doesn't fall back to default-src; the pane never submits a form.
    "form-action 'none'",
  ].join("; ");
}

module.exports = async (env = {}, argv = {}) => {
  const dev = argv.mode !== "production";
  const bench = Boolean(env.bench);
  const endpoint = (process.env.NYMFORM_ENDPOINT || "https://openrouter.ai/api/v1").replace(/\/+$/, "");
  const endpointOrigin = new URL(endpoint).origin;
  if (new URL(endpoint).protocol !== "https:") throw new Error("NYMFORM_ENDPOINT must use https");
  const config = {
    endpoint,
    defaultModel: process.env.NYMFORM_DEFAULT_MODEL || "openai/gpt-6-luna",
    version: pkg.version,
    commit: commitHash(),
  };

  let httpsOptions;
  if (dev && env.WEBPACK_SERVE) {
    // Only needed when serving; the dev certs are created by `npm run certs`.
    const devCerts = require("office-addin-dev-certs");
    httpsOptions = await devCerts.getHttpsServerOptions();
  }

  return {
    devtool: dev ? "source-map" : false,
    entry: { taskpane: "./src/taskpane/index.tsx", commands: "./src/taskpane/commands.ts" },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].[contenthash:8].js",
      chunkFilename: "[name].[contenthash:8].js",
      clean: true,
    },
    resolve: { extensions: [".ts", ".tsx", ".js"] },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: { loader: "ts-loader", options: { transpileOnly: true } },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        __NYMFORM_BENCH__: JSON.stringify(bench),
        __NYMFORM_CONFIG__: JSON.stringify(config),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
        csp: csp(endpointOrigin),
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/taskpane/commands.html",
        chunks: ["commands"],
        csp: csp(endpointOrigin),
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: "assets/*.png", to: "assets/[name][ext]" }],
      }),
      // build-info.json lets anyone match the served code to its source, and lets
      // scripts/check-release.mjs confirm a release was built with bench code off (T-P4).
      {
        apply(compiler) {
          compiler.hooks.thisCompilation.tap("NymformBuildInfo", (compilation) => {
            compilation.hooks.processAssets.tap(
              { name: "NymformBuildInfo", stage: webpack.Compilation.PROCESS_ASSETS_STAGE_ADDITIONAL },
              () => {
                const info = { name: "Nymform for Excel", ...config, bench, mode: dev ? "development" : "production" };
                compilation.emitAsset("build-info.json", new webpack.sources.RawSource(JSON.stringify(info, null, 2) + "\n"));
              },
            );
          });
        },
      },
    ],
    performance: { hints: false },
    devServer: {
      port: DEV_PORT,
      server: httpsOptions ? { type: "https", options: httpsOptions } : "https",
      headers: { "Access-Control-Allow-Origin": "*" },
      // No injected client: it would need a websocket the CSP does not allow. Reload by hand.
      client: false,
      hot: false,
      liveReload: false,
      static: false,
    },
  };
};

module.exports.DEV_URL = DEV_URL;
