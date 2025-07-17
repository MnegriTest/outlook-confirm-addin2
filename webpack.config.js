const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

module.exports = async (env, options) => {
    const dev = options.mode === "development";
    const config = {
        devtool: "source-map", // Genera source maps per un debugging più facile
        entry: {
            // Definisce i punti di ingresso per Webpack.
            // "commands" è per il codice che intercetta l'evento di invio.
            // "taskpane" è per il pannello laterale (se lo usi).
            // "dialog" è per il codice JavaScript del tuo popup di conferma.
            commands: "./src/commands/commands.js",
            taskpane: "./src/taskpane/taskpane.js",
            dialog: "./src/dialog/dialog.js", // <-- AGGIUNTA PER IL DIALOG
        },
        output: {
            clean: true, // Pulisce la cartella 'dist' prima di ogni build
        },
        resolve: {
            extensions: [".js", ".jsx", ".json", ".css", ".html"],
        },
        module: {
            rules: [
                {
                    test: /\.js$/,
                    exclude: /node_modules/,
                    use: {
                        loader: "babel-loader",
                        options: {
                            presets: ["@babel/preset-env"],
                        },
                    },
                },
                {
                    test: /\.html$/,
                    exclude: /node_modules/,
                    use: "html-loader",
                },
                {
                    test: /\.css$/,
                    exclude: /node_modules/,
                    use: ["style-loader", "css-loader"],
                },
                {
                    test: /\.(png|jpg|jpeg|gif|ico)$/,
                    type: "asset/resource",
                    generator: {
                        filename: "assets/[name][ext][query]", // Tutte le immagini andranno nella sottocartella 'assets'
                    },
                },
            ],
        },
        plugins: [
            new CleanWebpackPlugin(),
            new HtmlWebpackPlugin({
                filename: "commands.html",
                template: "./src/commands/commands.html",
                chunks: ["commands"], // Associa 'commands.html' all'entry point 'commands'
            }),
            new HtmlWebpackPlugin({
                filename: "taskpane.html",
                template: "./src/taskpane/taskpane.html",
                chunks: ["taskpane"], // Associa 'taskpane.html' all'entry point 'taskpane'
            }),
            new HtmlWebpackPlugin({ // <-- AGGIUNTA PER IL DIALOG
                filename: "dialog.html",
                template: "./src/dialog/dialog.html",
                chunks: ["dialog"], // Associa 'dialog.html' all'entry point 'dialog'
            }),
            new CopyWebpackPlugin({
                patterns: [
                    // Copia il manifest.xml direttamente nella root della cartella 'dist'
                    {
                        from: "manifest*.xml",
                        to: "[name][ext]",
                    },
                    // Copia tutti i file dalla cartella 'assets' sorgente alla sottocartella 'assets' in 'dist'
                    {
                        from: "assets/*",
                        to: "assets/[name][ext]",
                    },
                ],
            }),
            new webpack.ProvidePlugin({
                Promise: ["es6-promise", "Promise"],
            }),
        ],
        devServer: {
            headers: {
                "Access-Control-Allow-Origin": "*",
            },
            server: {
                type: "https", // Necessario per gli add-in di Office
                options: await devCerts.get // Genera e usa certificati di sviluppo per HTTPS
            },
            port: process.env.npm_package_config_dev_server_port || 3000, // Porta per il server di sviluppo
            hot: true, // Abilita Hot Module Replacement
        },
    };

    return config;
};