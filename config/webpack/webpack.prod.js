// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Specific environment; can be overriden.
process.env.NODE_ENV = process.env.NODE_ENV || 'prod';

const Merge = require('webpack-merge');
const CommonConfig = require('./webpack.common');
const DefinePlugin = require('webpack/lib/DefinePlugin');
const UglifyJSPlugin = require('uglifyjs-webpack-plugin');

module.exports = function (env) {
    return Merge(CommonConfig, {
        mode: 'production',
        plugins: [
            new DefinePlugin({
                'process.env': {
                    'NODE_ENV': JSON.stringify(process.env.NODE_ENV || 'prod'),
                }
            }),
            new UglifyJSPlugin({
                sourceMap: false,
            }),
        ],
    });
}
