// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

process.env.NODE_ENV = process.env.NODE_ENV || 'dev';

const Merge = require('webpack-merge');
const CommonConfig = require('./webpack.common');
const DefinePlugin = require('webpack/lib/DefinePlugin');

module.exports = function (env) {
    return Merge(CommonConfig, {
        mode: 'development',
        plugins: [
            new DefinePlugin({
                'process.env': {
                    'NODE_ENV': JSON.stringify(process.env.NODE_ENV || 'dev'),
                },
            }),
        ]
    });
}
