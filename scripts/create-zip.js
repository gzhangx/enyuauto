// Creates a smaller AWS Lambda deployment zip with only runtime artifacts.
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log('📦 Creating Lambda deployment package...\n');

const rootDir = process.cwd();
const zipName = path.join(rootDir, 'temp', 'lambda-deployment.zip');
const stagingDir = path.join(rootDir, 'temp', 'lambda-package');
const functionName = process.env.LAMBDA_FUNCTION_NAME || 'enyu-auto-v2-EnyuAutoFunction-r0jJLctkXxHF';
const awsRegion = process.env.AWS_REGION || 'us-east-1';

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

function copyProductionNodeModules(rootDir, stagingDir) {
  const packageJsonPath = path.join(rootDir, 'package.json');
  const lockJsonPath = path.join(rootDir, 'package-lock.json');
  if (!fs.existsSync(packageJsonPath)) {
    return;
  }

  const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
  const dependencyNames = new Set([
    ...Object.keys(packageJson.dependencies || {}),
    ...Object.keys(packageJson.optionalDependencies || {}),
  ]);

  if (fs.existsSync(lockJsonPath)) {
    const lockJson = JSON.parse(fs.readFileSync(lockJsonPath, 'utf8'));
    const packages = lockJson.packages || {};
    const copiedPackagePaths = new Set();

    for (const [pkgPath, meta] of Object.entries(packages)) {
      if (typeof pkgPath !== 'string' || !pkgPath.startsWith('node_modules/')) {
        continue;
      }
      if (meta?.dev || pkgPath.includes('node_modules/.bin/')) {
        continue;
      }
      const packagePath = path.join(rootDir, pkgPath);
      if (!fs.existsSync(packagePath)) {
        continue;
      }
      copyItem(packagePath, path.join(stagingDir, pkgPath));
      copiedPackagePaths.add(pkgPath);
    }

    for (const dependencyName of dependencyNames) {
      const dependencyPath = path.join(rootDir, 'node_modules', ...dependencyName.split('/'));
      const dependencyPkgPath = path.join('node_modules', ...dependencyName.split('/'));
      if (!fs.existsSync(dependencyPath) || copiedPackagePaths.has(dependencyPkgPath)) {
        continue;
      }
      copyItem(dependencyPath, path.join(stagingDir, dependencyPkgPath));
    }
    return;
  }

  for (const dependencyName of dependencyNames) {
    const dependencyPath = path.join(rootDir, 'node_modules', ...dependencyName.split('/'));
    const dependencyPkgPath = path.join('node_modules', ...dependencyName.split('/'));
    if (fs.existsSync(dependencyPath)) {
      copyItem(dependencyPath, path.join(stagingDir, dependencyPkgPath));
    }
  }
}

try {
  removeIfExists(zipName);
  removeIfExists(stagingDir);
  fs.mkdirSync(stagingDir, { recursive: true });
  console.log('  ✓ Cleaned old packaging output');

  console.log('  ⏳ Building latest Lambda code...');
  execSync('npm run build', { stdio: 'inherit', env: { ...process.env, AWS_PAGER: '' } });

  console.log('  ⏳ Copying runtime packages...');
  copyProductionNodeModules(rootDir, stagingDir);

  const runtimeEntries = ['dist', 'package.json', 'package-lock.json'];
  for (const entry of runtimeEntries) {
    const sourcePath = path.join(rootDir, entry);
    if (fs.existsSync(sourcePath)) {
      copyItem(sourcePath, path.join(stagingDir, entry));
    }
  }

  console.log('  ⏳ Compressing runtime files...');
  const command = `powershell -NoProfile -NonInteractive -Command "Compress-Archive -Path '${stagingDir}\\*' -DestinationPath '${zipName}' -Force"`;
  execSync(command, { stdio: 'inherit', env: { ...process.env, AWS_PAGER: '' } });

  const stats = fs.statSync(zipName);
  const fileSizeMB = (stats.size / (1024 * 1024)).toFixed(2);

  console.log(`\n✅ Deployment package created: ${zipName}`);
  console.log(`   Size: ${fileSizeMB} MB\n`);
  console.log('📤 Uploading package to AWS Lambda...');
  execSync(`aws lambda update-function-code --function-name ${functionName} --region ${awsRegion} --zip-file fileb://${zipName} --no-cli-pager`, { stdio: 'inherit', env: { ...process.env, AWS_PAGER: '' } });
  console.log(`\n✅ Lambda function updated: ${functionName}`);
  console.log('📤 This archive now contains only production runtime artifacts:');
  console.log('   - dist/');
  console.log('   - package.json');
  console.log('   - node_modules/ (pruned for production)');
  console.log('');
} catch (error) {
  console.error('\n❌ Error creating zip file:', error.message);
  process.exit(1);
}
