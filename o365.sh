#!/bin/sh

LIB=Microsoft.Office365.ClientLib.JS
VER=1.0.22
O365=./$LIB.$VER/content/services/office365/scripts/

# Install lib
# mono nuget.exe $LIB

# Write header
echo "/// $LIB.$VER" > o365.js
echo "/// patched by masataka takeuchi for nodejs support" >> o365.js

# Polyfill window, XMLHttpRequest
echo 'var window = require("./lib/node-window"), XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;' >> o365.js

# Concat for node module
cat $O365/utility.js $O365/o365auth.js $O365/o365discovery.js $O365/exchange.js $O365/sharepoint.js >> o365.js
echo 'module.exports = { Microsoft: Microsoft, O365Auth: O365Auth, O365Discovery: O365Discovery };' >> o365.js

# Support Buffer
sed -e "s/\(typeof request.data === 'string'\)/(\1) || (Buffer.isBuffer(request.data))/" \
    -e "s/\(var xhr =.*;\)/\1 xhr.responseType = request.responseType;/" \
    -e "s/\(Request(this.getPath(\"content\"));\)/\1 request.responseType = 'buffer';/" \
    -e "s/responseText/response/" \
    -e "s/JSON.stringify({ \"contentStream\": contentStream })/contentStream/" \
  o365.js > o365.js.tmp && mv o365.js.tmp o365.js

# Implement Item Fetcher cast
L1="ItemFetcher.prototype.asFile = function () { return new FileFetcher(this.context, this.path); };"
L2="ItemFetcher.prototype.asFolder = function () { return new FolderFetcher(this.context, this.path); };"
sed -e "s/\(FileServices.ItemFetcher = ItemFetcher;\)/$L1 $L2 \1/" \
  o365.js > o365.js.tmp && mv o365.js.tmp o365.js

