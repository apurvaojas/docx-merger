const path = require('path');

module.exports = {
    devtool: 'source-map',
    mode: 'production',
    entry: path.join(__dirname, "src/index.ts"),
    output: {
        filename: 'docx-merger.js',
        path: path.join(__dirname, 'dist'),
        libraryTarget: 'umd',
        library: 'DocxMerger',
    },
    module: {
        rules: [
            {
                test: /\.ts$/,
                exclude: /node_modules/,
                loader: 'babel-loader',
                options: {
                    presets: ['@babel/preset-env','@babel/preset-typescript']
                }
            }
        ]
    },
    plugins: []
};
