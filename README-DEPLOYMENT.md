# AWS Deployment Guide

Updated deployment:
sam build
sam deploy

## Prerequisites

### 1. Install AWS CLI
```bash
# Windows (using chocolatey)
choco install awscli

# Or download from: https://aws.amazon.com/cli/
```

### 2. Install AWS SAM CLI
```bash
# Windows (using chocolatey)
choco install aws-sam-cli

# Or download from: https://docs.aws.amazon.com/serverless-application-model/latest/developerguide/install-sam-cli.html
```

### 3. Configure AWS Credentials
```bash
aws configure
```
You'll need:
- AWS Access Key ID
- AWS Secret Access Key
- Default region (e.g., us-east-1)
- Output format (json)

## Deployment Methods

### Option 1: Using the Deployment Script (Recommended)

**Windows:**
```bash
deploy.bat
```

**Linux/Mac:**
```bash
chmod +x deploy.sh
./deploy.sh
```

### Option 2: Using npm Scripts

**First time (guided):**
```bash
npm run deploy:guided
```

**Subsequent deployments:**
```bash
npm run deploy:quick
```

### Option 3: Manual SAM Commands

**Build:**
```bash
npm run build
sam build
```

**Deploy (first time):**
```bash
sam deploy --guided
```

**Deploy (subsequent):**
```bash
sam deploy
```

## First-Time Deployment

When deploying for the first time with `--guided` flag, you'll be prompted for:

1. **Stack Name**: e.g., `enyu-auto-stack`
2. **AWS Region**: e.g., `us-east-1`
3. **Confirm changes before deploy**: Y/n
4. **Allow SAM CLI IAM role creation**: Y
5. **Disable rollback**: n
6. **Save arguments to configuration file**: Y
7. **Configuration file name**: [samconfig.toml]
8. **Configuration environment**: [default]

These settings will be saved in `samconfig.toml` for future deployments.

## Updating the Lambda Function

After the first deployment, any code changes can be deployed with:

```bash
npm run deploy:quick
```

Or on Windows:
```bash
deploy.bat
```
Then select option 2 for quick deployment.

## Getting Your API Endpoint

After deployment, get your API endpoint URL:

```bash
aws cloudformation describe-stacks \
  --stack-name enyu-auto-stack \
  --query "Stacks[0].Outputs[?OutputKey=='EnyuAutoApi'].OutputValue" \
  --output text
```

Or check the AWS CloudFormation Console:
https://console.aws.amazon.com/cloudformation

## Testing Your Deployed Function

```bash
curl "https://YOUR-API-ID.execute-api.us-east-1.amazonaws.com/Prod/process?line=5&doit=true"
```

## Deleting the Stack

To remove all AWS resources:

```bash
sam delete --stack-name enyu-auto-stack
```

## Troubleshooting

### Build Fails
- Make sure TypeScript compiles: `npm run build`
- Check that all dependencies are installed: `npm install`

### Deployment Fails
- Verify AWS credentials: `aws sts get-caller-identity`
- Check IAM permissions (need CloudFormation, Lambda, API Gateway, IAM)
- Ensure S3 bucket name is unique (if manually specified)

### Function Errors
- Check CloudWatch Logs:
  ```bash
  sam logs -n EnyuAutoFunction --stack-name enyu-auto-stack --tail
  ```

### Need to Update Configuration
- Edit `samconfig.toml` or run `sam deploy --guided` again

## CI/CD Integration

For GitHub Actions, create `.github/workflows/deploy.yml`:

```yaml
name: Deploy to AWS

on:
  push:
    branches: [ main ]

jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-node@v3
        with:
          node-version: '20'
      - uses: aws-actions/setup-sam@v2
      - uses: aws-actions/configure-aws-credentials@v2
        with:
          aws-access-key-id: ${{ secrets.AWS_ACCESS_KEY_ID }}
          aws-secret-access-key: ${{ secrets.AWS_SECRET_ACCESS_KEY }}
          aws-region: us-east-1
      - run: npm ci
      - run: npm run build
      - run: sam build
      - run: sam deploy --no-confirm-changeset --no-fail-on-empty-changeset
```

## Environment Variables

To add environment variables to your Lambda function, edit `template.yaml`:

```yaml
Environment:
  Variables:
    NODE_ENV: production
    MY_SECRET: !Sub '{{resolve:secretsmanager:my-secret-name}}'
```

## Secrets Management

For sensitive data like `enyu_secs.json`:

1. Create a secret in AWS Secrets Manager
2. Grant Lambda function permission to access it
3. Retrieve in code:
   ```javascript
   const AWS = require('aws-sdk');
   const secretsManager = new AWS.SecretsManager();
   const secret = await secretsManager.getSecretValue({ SecretId: 'my-secret' }).promise();
   ```
