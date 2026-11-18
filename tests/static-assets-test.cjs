/* Community Connect - static asset path test (zero dependencies).
   Verifies that the files referenced by the public HTML, CSS, JavaScript,
   and service worker are present in the repository.

   Run with: node tests/static-assets-test.cjs */

var fs = require("fs");
var path = require("path");

var ROOT = path.resolve(__dirname, "..");
var failures = [];

function fail(message) {
  failures.push(message);
}

function read(relativePath) {
  return fs.readFileSync(path.join(ROOT, relativePath), "utf8");
}

function exists(relativePath) {
  var clean = relativePath.split("?")[0].split("#")[0];
  return !!clean && fs.existsSync(path.join(ROOT, clean));
}

function unique(values) {
  return values.filter(function (value, index) {
    return values.indexOf(value) === index;
  });
}

// Check local src/href assets in every top-level HTML page.
var htmlFiles = fs.readdirSync(ROOT).filter(function (name) {
  return /\.html$/i.test(name);
});

htmlFiles.forEach(function (name) {
  var html = read(name);
  var pattern = /(?:src|href)=["']([^"']+)["']/gi;
  var match;
  while ((match = pattern.exec(html))) {
    var asset = match[1];
    if (/^(?:https?:|data:|blob:|#|mailto:|tel:|\/\/)/i.test(asset)) continue;
    if (!/\.(?:html|css|js|json|png|jpe?g|gif|webp|svg|ico|webmanifest)$/i.test(asset)) continue;
    if (!exists(asset)) fail(name + " references missing asset: " + asset);
  }
});

// Check bundled image paths used by the JavaScript seed data and helpers.
var script = read("js/script.js");
var imagePattern = /["'](images\/[^"']+)["']/g;
var imagePaths = [];
var imageMatch;
while ((imageMatch = imagePattern.exec(script))) imagePaths.push(imageMatch[1]);
imagePaths = unique(imagePaths);
imagePaths.forEach(function (asset) {
  if (!exists(asset)) fail("js/script.js references missing image: " + asset);
});

// Check CSS url(...) assets.
var css = read("css/style.css");
var cssPattern = /url\(["']?([^"')]+)["']?\)/g;
var cssMatch;
while ((cssMatch = cssPattern.exec(css))) {
  var cssAsset = cssMatch[1];
  if (/^(?:https?:|data:|blob:|#)/i.test(cssAsset)) continue;
  var cssPath = cssAsset.replace(/^\.\.\//, "");
  if (!exists(cssPath)) fail("css/style.css references missing asset: " + cssAsset);
}

// Check every service-worker core file.
var serviceWorker = read("sw.js");
var corePattern = /["']((?:[\w.-]+\/)*[\w.-]+\.(?:html|css|js|json|png|jpe?g|gif|webp|svg))["']/g;
var coreMatch;
while ((coreMatch = corePattern.exec(serviceWorker))) {
  if (!exists(coreMatch[1])) fail("sw.js references missing asset: " + coreMatch[1]);
}

// Bundled images should avoid spaces and legacy .jfif extensions so they
// have the same URL on GitHub Pages, Node, and file:// launches.
var imagesDir = path.join(ROOT, "images");
fs.readdirSync(imagesDir).forEach(function (name) {
  if (!/\.(?:png|jpe?g|gif|webp|svg|jfif)$/i.test(name)) return;
  if (/\s/.test(name) || /\.jfif$/i.test(name)) {
    fail("images/" + name + " must use a lowercase, space-free .jpg/.webp path");
  }
});

if (failures.length) {
  console.error("STATIC ASSET TEST: " + failures.length + " failure(s)");
  failures.forEach(function (message) { console.error("FAIL  " + message); });
  process.exit(1);
}

console.log("STATIC ASSET TEST: ALL PASS (" + htmlFiles.length + " HTML pages, " +
  imagePaths.length + " JavaScript image paths)");
