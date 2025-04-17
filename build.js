"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
// @ts-ignore
var fs = require("fs");
// @ts-ignore
var path = require("path");
var directoryPath = path.resolve("./", 'src');
var outputFilePath = './build/SlsBns.osts';
var combinedContent = "";
// Read all files in the directory
fs.readdir(directoryPath, function (err, files) {
    if (err) {
        console.error('Error reading directory:', err);
        return;
    }
    // Filter to include only files (not directories)
    var allFiles = files.filter(function (file) {
        var filePath = path.join(directoryPath, file);
        return fs.statSync(filePath).isFile();
    });
    allFiles.forEach(function (file) {
        var filePath = path.resolve("./src/", "".concat(file));
        var fileContent = fs.readFileSync(filePath, 'utf-8');
        combinedContent += fileContent + '\n\n';
    });
    // Write to the output file
    fs.writeFileSync(outputFilePath, combinedContent);
    console.log("Files combined into ".concat(outputFilePath));
});
