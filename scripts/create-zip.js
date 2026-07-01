// Creates a smaller AWS Lambda deployment zip with only runtime artifacts.
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log('📦 Creating Lambda deployment package...\n');

const rootDir = process.cwd();
const zipName = path.join(rootDir, 'temp', 'lambda-deployment.zip');
const stagingDir = path.join(rootDir, 'temp', 'lambda-package');

function removeIfExists(targetPath) {
  if (fs.existsSync(targetPath)) {
    fs.rmSync(targetPath, { recursive: true, force: true });
  }
}

function copyItem(sourcePath, targetPath) {
  const stats = fs.statSync(sourcePath);
  if (stats.isDirectory()) {
    fs.mkdirSync(targetPath, { recursive: true });
    for (const entry of fs.readdirSync(sourcePath)) {
      copyItem(path.join(sourcePath, entry), path.join(targetPath, entry));
    }
    return;
  }

  fs.mkdirSync(path.dirname(targetPath), { recursive: true });
  fs.copyFileSync(sourcePath, targetPath);
}

try {
  removeIfExists(zipName);
  removeIfExists(stagingDir);
  fs.mkdirSync(stagingDir, { recursive: true });
  console.log('  ✓ Cleaned old packaging output');

  console.log('  ⏳ Pruning dev-only dependencies...');
  execSync('npm prune --omit=dev', { stdio: 'inherit' });

  const runtimeEntries = ['dist', 'package.json', 'package-lock.json', 'node_modules'];
  for (const entry of runtimeEntries) {
    const sourcePath = path.join(rootDir, entry);
    if (fs.existsSync(sourcePath)) {
      copyItem(sourcePath, path.join(stagingDir, entry));
    }
  }

  console.log('  ⏳ Compressing runtime files...');
  const command = `powershell Compress-Archive -Path "${stagingDir}\\*" -DestinationPath "${zipName}" -Force`;
  execSync(command, { stdio: 'inherit' });

  const stats = fs.statSync(zipName);
  const fileSizeMB = (stats.size / (1024 * 1024)).toFixed(2);

  console.log(`\n✅ Deployment package created: ${zipName}`);
  console.log(`   Size: ${fileSizeMB} MB\n`);
  console.log('📤 This archive now contains only production runtime artifacts:');
  console.log('   - dist/');
  console.log('   - package.json');
  console.log('   - node_modules/ (pruned for production)');
  console.log('');
} catch (error) {
  console.error('\n❌ Error creating zip file:', error.message);
  process.exit(1);
}
