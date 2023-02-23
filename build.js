const { build } = require("esbuild");
const { dependencies } = require("./package.json");

const entryFile = "src/index.js";
const shared = {
  bundle: true,
  entryPoints: [entryFile],
  // Treat all dependencies in package.json as externals to keep bundle size to a minimum
  external: Object.keys(dependencies),
  logLevel: "info",
  minify: true,
  sourcemap: true,
};

build({
  ...shared,
  format: "esm",
  outfile: "./dist/index.esm.js",
  target: ["es2018", "node12.22.0"],
});

build({
  ...shared,
  format: "cjs",
  outfile: "./dist/index.cjs.js",
  target: ["es2018", "node12.22.0"],
});
