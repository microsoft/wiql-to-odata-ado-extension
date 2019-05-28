// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const path = require('path');
const webpack = require('webpack');
const CssExtractPlugin = require('mini-css-extract-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');

const rootDir = path.resolve(__dirname, '../..');

if (process.env.ENV_FILE == null) {
    process.env.ENV_FILE = process.env.NODE_ENV;
}

const extractSass = new CssExtractPlugin({
    filename: '[name].[hash].css',
});

const indexHtml = new HtmlWebpackPlugin({
    template: path.join(rootDir, './src/index.html'),
    filename: path.join(rootDir, './dist/index.html'),
    inject: false,
});

const dialogHtml = new HtmlWebpackPlugin({
    template: path.join(rootDir, './src/dialog.html'),
    filename: path.join(rootDir, './dist/dialog.html'),
    inject: false,
});

module.exports = {
    target: 'web',
    entry: {
        app: './src/app/app.ts',
        dialog: './src/dialog/dialog.tsx',
        polyfills: './src/polyfills.ts'
    },
    output: {
        path: path.join(rootDir, './dist/bundles'),
        filename: '[name].[hash].js',
        libraryTarget: 'amd'
    },
    externals: [
        /^TFS\/.*/,
        /^VSS\/.*/,
    ],
    resolve: {
        extensions: [
            '.ts',
            '.js',
            ".tsx",
        ],
        alias: {
            '@ms/configuration-environment': path.join(rootDir, './config/environment/environment.' + process.env.ENV_FILE),
        }
    },
    module: {
        rules: [
            {
                test: /\.tsx?$/,
                loader: 'ts-loader',
                options: {

                }
            },
            {
                enforce: 'pre',
                test: /\.js$/,
                loader: 'source-map-loader',
                exclude: [
                    /\/node_modules\//
                ]
            },
            {
                test: /\.scss$|\.sass$/,
                use: [
                    {
                        loader: CssExtractPlugin.loader,
                        options: { }
                    },
                    "css-loader",
                    "sass-loader"
                ],
            }
        ]
    },
    devtool: 'source-map',
    plugins: [
        extractSass,
        indexHtml,
        dialogHtml,
    ],
}
