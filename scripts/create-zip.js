// no longer used, pushing to github
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log('📦 Creating Lambda deployment package...\n');

const zipName = 'temp\\lambda-deployment.zip';

// Remove old zip if exists
if (fs.existsSync(zipName)) {
  fs.unlinkSync(zipName);
  console.log('  ✓ Removed old deployment package');
}

// Create zip using PowerShell (Windows)
try {
  console.log('  ⏳ Compressing files...');
  
  // Use PowerShell to create zip
  const command = `powershell Compress-Archive -Path dist,node_modules,package.json,enyu_secs.json,index.mjs -DestinationPath ${zipName} -Force`;
  execSync(command, { stdio: 'inherit' });
  
  const stats = fs.statSync(zipName);
  const fileSizeMB = (stats.size / (1024 * 1024)).toFixed(2);
  
  console.log(`\n✅ Deployment package created: ${zipName}`);
  console.log(`   Size: ${fileSizeMB} MB\n`);
  
  console.log('📤 Manual Deployment Steps:\n');
  console.log('1. Go to AWS Lambda Console:');
  console.log('   https://console.aws.amazon.com/lambda/home\n');
  console.log('2. Click "Create function"');
  console.log('   - Function name: enyu-auto');
  console.log('   - Runtime: Node.js 20.x');
  console.log('   - Click "Create function"\n');
  console.log('3. In the "Code" tab:');
  console.log('   - Click "Upload from" → ".zip file"');
  console.log(`   - Upload: ${zipName}`);
  console.log('   - Click "Save"\n');
  console.log('4. In "Configuration" → "General configuration":');
  console.log('   - Click "Edit"');
  console.log('   - Handler: index.handler');
  console.log('   - Timeout: 5 minutes (300 seconds)');
  console.log('   - Memory: 512 MB');
  console.log('   - Click "Save"\n');
  console.log('5. Add API Gateway trigger:');
  console.log('   - Click "Add trigger"');
  console.log('   - Select "API Gateway"');
  console.log('   - Create new API (HTTP API)');
  console.log('   - Security: Open');
  console.log('   - Click "Add"\n');
  console.log('6. Test your function with the API endpoint!\n');
  
} catch (error) {
  console.error('\n❌ Error creating zip file:', error.message);
  console.log('\nAlternative: Create zip manually:');
  console.log('1. Select these folders/files: dist, node_modules, package.json, enyu_secs.json');
  console.log('2. Right-click → Send to → Compressed (zipped) folder');
  console.log('3. Rename to lambda-deployment.zip');
  process.exit(1);
}
