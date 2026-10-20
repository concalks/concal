/* eslint-disable no-undef */

const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const {readFileSync} = require("node:fs");
const {resolve} = require("node:path");

const urlDev = "https://localhost:3000/";
const urlProd = "https://concalks.github.io/concal/"; //TODO CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

module.exports = async (env, options) => {
    const dev = options.mode === "development";
    let httpsOptions;
    if (dev) {
        httpsOptions = {
            key: readFileSync(resolve(__dirname, "localhost-key.pem")),
            cert: readFileSync(resolve(__dirname, "localhost.pem")),
        };
    }
    const config = {
        devtool: "source-map",
        entry: {
            polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
            taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane_appointment_creation.html"],
            commands: "./src/commands/commands.ts",
        },
        output: {
            clean: true,
        },
        resolve: {
            extensions: [".ts", ".html", ".js"],
        },
        module: {
            rules: [
                {
                    test: /\.ts$/,
                    exclude: /node_modules/,
                    use: {
                        loader: "babel-loader"
                    },
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
            new HtmlWebpackPlugin({
                filename: "taskpane_appointment_creation.html",
                template: "./src/taskpane/taskpane_appointment_creation.html",
                chunks: ["polyfill", "taskpane"],
            }),
            new CopyWebpackPlugin({
                patterns: [
                    {
                        from: "assets/*",
                        to: "assets/[name][ext][query]",
                    },
                    {
                        from: "manifest*.xml",
                        to: "[name]" + "[ext]",
                        transform(content) {
                            if (dev) {
                                return content;
                            } else {
                                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
                            }
                        },
                    },
                ],
            }),
            new HtmlWebpackPlugin({
                filename: "commands.html",
                template: "./src/commands/commands.html",
                chunks: ["polyfill", "commands"],
            }),
        ],
        devServer: {
            headers: {
                "Access-Control-Allow-Origin": "*",
            },
            server: {
                type: "https",
                options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : httpsOptions,
            },
            port: process.env.npm_package_config_dev_server_port || 3000,
        },
    };

    return config;
};
