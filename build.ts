// @ts-ignore
import * as fs from 'fs';
// @ts-ignore
import * as path from 'path';

const directoryPath = path.resolve("./", 'src');
const outputFilePath = './build/SlsBns.osts';
let combinedContent = "";

// Read all files in the directory
fs.readdir(directoryPath, (err: Error, files: Array<File>) => {
    if (err) {
        console.error('Error reading directory:', err);
        return;
    }

    // Filter to include only files (not directories)
    const allFiles = files.filter(file => {
        const filePath = path.join(directoryPath, file);
        return fs.statSync(filePath).isFile();
    });

    allFiles.forEach(file => {
        const filePath = path.resolve("./src/", `${file}`);
        const fileContent = fs.readFileSync(filePath, 'utf-8');
        combinedContent += fileContent + '\n\n';
    });

    // Write to the output file
    fs.writeFileSync(outputFilePath, combinedContent);

    console.log(`Files combined into ${outputFilePath}`);
});