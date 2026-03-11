const fs = require('fs');
const path = require('path');

const pkg = require('../package.json');
const version = pkg.version;

const configPath = path.join(__dirname, 'target.json');
const configData = JSON.parse(fs.readFileSync(configPath, 'utf8'));
configData.version = version;

const buildDir = path.join(__dirname, '..', 'build');
if (!fs.existsSync(buildDir)) {
  fs.mkdirSync(buildDir, { recursive: true });
}

targetPath = path.join(buildDir, 'config.json'),

fs.writeFileSync(targetPath, JSON.stringify(configData, null, 2));
console.log(`Generated ${targetPath} with version ${version}`);
