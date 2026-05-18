# Deployment guide — Startup Scout on AWS

This document describes how to move Startup Scout from a local Mac to AWS.
See the architecture proposal in the handoff notes for the rationale behind
each service choice.

## What you're working with

| Thing | Detail |
|---|---|
| Language / framework | Python 3.11, Flask, Gunicorn |
| Container | `Dockerfile` at repo root — already production-ready |
| Persistent data | `Startup database.xlsx`, `users.json`, `collections.json`, `startups/` — all under `DATA_DIR` |
| Auth | Shared password fallback **+** full per-user email/password system in `users.py` |
| AI features | Anthropic API (`/api/ai/*` routes) — needs `ANTHROPIC_API_KEY` |
| Health check | `GET /health` → `{"ok": true}` |
| Required env vars | See `.env.example` — every variable is documented there |

---

## Phase 1 — Get it running on AWS (no code changes)

### 1. Prerequisites

- AWS CLI configured (`aws configure` or IAM Identity Center)
- Docker installed locally
- The repo pushed to GitHub (needed for CI/CD in step 6)

### 2. Create ECR repository

```bash
aws ecr create-repository \
  --repository-name startup-scout \
  --region eu-west-1
```

Push a first image manually to confirm it works:

```bash
AWS_ACCOUNT=$(aws sts get-caller-identity --query Account --output text)
REGION=eu-west-1

aws ecr get-login-password --region $REGION \
  | docker login --username AWS \
    --password-stdin $AWS_ACCOUNT.dkr.ecr.$REGION.amazonaws.com

docker build -t startup-scout .
docker tag startup-scout:latest \
  $AWS_ACCOUNT.dkr.ecr.$REGION.amazonaws.com/startup-scout:latest
docker push \
  $AWS_ACCOUNT.dkr.ecr.$REGION.amazonaws.com/startup-scout:latest
```

### 3. Create the VPC and networking

Use the AWS console wizard ("Launch VPC wizard") or Terraform.
You need:

- 1 VPC (`10.0.0.0/16`)
- 2 public subnets (ALB lives here; spread across 2 AZs for redundancy)
- 2 private subnets (Fargate tasks live here)
- Internet Gateway + NAT Gateway (so private tasks can pull ECR images)
- Security groups:
  - `alb-sg`: inbound 443 from `0.0.0.0/0`; outbound 8080 to `app-sg`
  - `app-sg`: inbound 8080 from `alb-sg`; outbound 443 to `0.0.0.0/0` (ECR, Secrets, SES)
  - `efs-sg`: inbound 2049 (NFS) from `app-sg`

### 4. Create EFS and copy data

```bash
# Create the file system
aws efs create-file-system \
  --performance-mode generalPurpose \
  --tags Key=Name,Value=startup-scout-data

# Note the FileSystemId, then create mount targets in each private subnet:
aws efs create-mount-target \
  --file-system-id fs-XXXXXXXX \
  --subnet-id subnet-XXXXXXXX \
  --security-groups sg-XXXXXXXX   # efs-sg
```

Copy existing data to EFS (easiest via a temporary EC2 in the same VPC):

```bash
# On the EC2:
sudo mount -t efs fs-XXXXXXXX:/ /mnt/data
sudo cp "Startup database.xlsx" /mnt/data/
sudo cp users.json collections.json /mnt/data/   # if they exist
```

> **Important:** EFS uses NFS. The app opens the Excel file with openpyxl which
> takes an exclusive write lock. This is fine for a small team but you may see
> occasional "file locked" errors if several people save at exactly the same time.
> Phase 2 (database migration) eliminates this.

### 5. Store secrets

```bash
aws secretsmanager create-secret \
  --name startup-scout/env \
  --secret-string '{
    "SECRET_KEY": "your-64-hex-chars",
    "APP_PASSWORD": "your-strong-password",
    "DATA_DIR": "/data",
    "PORT": "8080",
    "ANTHROPIC_API_KEY": "sk-ant-...",
    "SMTP_HOST": "email-smtp.eu-west-1.amazonaws.com",
    "SMTP_PORT": "587",
    "SMTP_USER": "...",
    "SMTP_PASS": "...",
    "SMTP_FROM": "startupscout@yourdomain.com"
  }'
```

### 6. Create ECS cluster, task definition and service

**Cluster:**
```bash
aws ecs create-cluster --cluster-name startup-scout-cluster
```

**Task definition** (save as `task-definition.json` and register it):
```json
{
  "family": "startup-scout",
  "networkMode": "awsvpc",
  "requiresCompatibilities": ["FARGATE"],
  "cpu": "512",
  "memory": "1024",
  "executionRoleArn": "arn:aws:iam::ACCOUNT:role/ecsTaskExecutionRole",
  "taskRoleArn":      "arn:aws:iam::ACCOUNT:role/ecsTaskRole",
  "containerDefinitions": [{
    "name": "startup-scout",
    "image": "ACCOUNT.dkr.ecr.REGION.amazonaws.com/startup-scout:latest",
    "portMappings": [{ "containerPort": 8080 }],
    "secrets": [
      { "name": "SECRET_KEY",        "valueFrom": "arn:aws:secretsmanager:...:startup-scout/env:SECRET_KEY::" },
      { "name": "APP_PASSWORD",      "valueFrom": "arn:aws:secretsmanager:...:startup-scout/env:APP_PASSWORD::" },
      { "name": "DATA_DIR",          "valueFrom": "arn:aws:secretsmanager:...:startup-scout/env:DATA_DIR::" },
      { "name": "ANTHROPIC_API_KEY", "valueFrom": "arn:aws:secretsmanager:...:startup-scout/env:ANTHROPIC_API_KEY::" },
      { "name": "SMTP_HOST",         "valueFrom": "arn:aws:secretsmanager:...:startup-scout/env:SMTP_HOST::" },
      { "name": "SMTP_PORT",         "valueFrom": "arn:aws:secretsmanager:...:startup-scout/env:SMTP_PORT::" },
      { "name": "SMTP_USER",         "valueFrom": "arn:aws:secretsmanager:...:startup-scout/env:SMTP_USER::" },
      { "name": "SMTP_PASS",         "valueFrom": "arn:aws:secretsmanager:...:startup-scout/env:SMTP_PASS::" },
      { "name": "SMTP_FROM",         "valueFrom": "arn:aws:secretsmanager:...:startup-scout/env:SMTP_FROM::" }
    ],
    "mountPoints": [{
      "sourceVolume": "efs-data",
      "containerPath": "/data"
    }],
    "healthCheck": {
      "command": ["CMD-SHELL", "curl -f http://localhost:8080/health || exit 1"],
      "interval": 30, "timeout": 5, "retries": 3, "startPeriod": 10
    },
    "logConfiguration": {
      "logDriver": "awslogs",
      "options": {
        "awslogs-group": "/ecs/startup-scout",
        "awslogs-region": "eu-west-1",
        "awslogs-stream-prefix": "ecs"
      }
    }
  }],
  "volumes": [{
    "name": "efs-data",
    "efsVolumeConfiguration": {
      "fileSystemId": "fs-XXXXXXXX",
      "rootDirectory": "/",
      "transitEncryptionPort": 2049
    }
  }]
}
```

```bash
aws ecs register-task-definition --cli-input-json file://task-definition.json
```

**Service:**
```bash
aws ecs create-service \
  --cluster startup-scout-cluster \
  --service-name startup-scout-service \
  --task-definition startup-scout \
  --desired-count 1 \
  --launch-type FARGATE \
  --network-configuration "awsvpcConfiguration={
    subnets=[subnet-PRIVATE1,subnet-PRIVATE2],
    securityGroups=[sg-APP],
    assignPublicIp=DISABLED
  }" \
  --load-balancers "targetGroupArn=arn:...,containerName=startup-scout,containerPort=8080" \
  --deployment-configuration "minimumHealthyPercent=100,maximumPercent=200"
```

### 7. ALB, ACM certificate and Route 53

1. **ALB** — create in the two public subnets, attach `alb-sg`
2. **Target group** — type `ip`, port 8080, health check path `/health`
3. **ACM** — request a cert for your domain (DNS validation, takes ~2 min)
4. **ALB listener** — HTTPS :443, forward to target group, attach ACM cert
5. **Route 53** — create an A record (alias) pointing at the ALB DNS name

### 8. Set up GitHub Actions

In your GitHub repo, go to **Settings → Secrets and variables → Actions** and add:

| Secret | Value |
|---|---|
| `AWS_REGION` | `eu-west-1` |
| `AWS_ACCOUNT_ID` | your 12-digit account ID |
| `ECR_REPOSITORY` | `startup-scout` |
| `ECS_CLUSTER` | `startup-scout-cluster` |
| `ECS_SERVICE` | `startup-scout-service` |
| `ECS_CONTAINER` | `startup-scout` |

The workflow uses **OIDC** (no long-lived AWS keys stored in GitHub).
Create the IAM role it assumes:

```bash
# Trust policy — allows GitHub Actions to assume this role
cat > trust.json << 'EOF'
{
  "Version": "2012-10-17",
  "Statement": [{
    "Effect": "Allow",
    "Principal": { "Federated": "arn:aws:iam::ACCOUNT:oidc-provider/token.actions.githubusercontent.com" },
    "Action": "sts:AssumeRoleWithWebIdentity",
    "Condition": {
      "StringLike": {
        "token.actions.githubusercontent.com:sub": "repo:YOUR_ORG/YOUR_REPO:ref:refs/heads/main"
      }
    }
  }]
}
EOF

aws iam create-role \
  --role-name github-actions-startup-scout \
  --assume-role-policy-document file://trust.json

# Attach the permissions the workflow needs
aws iam attach-role-policy \
  --role-name github-actions-startup-scout \
  --policy-arn arn:aws:iam::aws:policy/AmazonEC2ContainerRegistryPowerUser

aws iam attach-role-policy \
  --role-name github-actions-startup-scout \
  --policy-arn arn:aws:iam::aws:policy/AmazonECS_FullAccess
```

Push to `main` — the workflow in `.github/workflows/deploy.yml` builds the image,
pushes it to ECR, updates the task definition and triggers a rolling deploy.
The step waits for the service to stabilise before marking the run green.

### 9. SES — verify your sender domain

```bash
aws ses verify-domain-identity --domain yourdomain.com --region eu-west-1
# Follow the DNS verification instructions in the console
# Then request production access (SES starts in sandbox mode)
```

Set `SMTP_HOST`, `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM` in Secrets Manager
(step 5). The app's email verification and password-reset flows will work
automatically once those are set.

---

## Phase 2 — Migrate to PostgreSQL (when ready)

The single-shared-Excel-file model works but concurrent writes can conflict.
When the team grows beyond ~5 active users, migrate to RDS.

**Checklist:**
- [ ] Provision RDS PostgreSQL `db.t4g.micro` in private subnets
- [ ] Add `DATABASE_URL` to Secrets Manager
- [ ] Add `psycopg2-binary` and `SQLAlchemy` to `requirements.txt`
- [ ] Write schema: `startups`, `notes`, `collections`, `collection_members`, `users`
- [ ] Write a one-time migration script that reads `Startup database.xlsx` and
      inserts rows into PostgreSQL
- [ ] Replace the openpyxl calls in `app.py` with SQLAlchemy queries
      (all writes are in clearly-named functions: `load_startups`,
      `api_add_startup`, `api_edit_startup`, `api_add_update`, etc.)
- [ ] Drop ALB sticky sessions (no longer needed without file-based sessions)
- [ ] At this point you can switch to App Runner instead of ECS+ALB, saving ~$20/mo

---

## Phase 3 — Enable per-user accounts (hours of work, already coded)

`users.py` has a complete user management system — create, verify, authenticate,
password reset, roles, active/disabled. It's dormant because the current deploy
only uses the shared password. To activate it:

- [ ] Verify SES is working (send a test email via `email_utils.py`)
- [ ] Decide whether to keep the shared-password fallback (`APP_PASSWORD`) or remove it
- [ ] Create accounts for each team member (use the `/register` route or a seeding script)
- [ ] Remove or rotate `APP_PASSWORD` once everyone has individual accounts

---

## Runbook — day-to-day operations

### Deploy a new version
Push to `main`. GitHub Actions handles everything. Takes ~3 minutes.

### Roll back
```bash
# Find the previous task definition revision
aws ecs describe-task-definition --task-definition startup-scout:PREV_REVISION

# Point the service at it
aws ecs update-service \
  --cluster startup-scout-cluster \
  --service startup-scout-service \
  --task-definition startup-scout:PREV_REVISION
```

### Tail application logs
```bash
aws logs tail /ecs/startup-scout --follow
```

### Update a secret (e.g. rotate the app password)
```bash
aws secretsmanager update-secret \
  --secret-id startup-scout/env \
  --secret-string '{"APP_PASSWORD":"new-password", ...all other keys...}'

# Force ECS to pick up the new value by restarting tasks
aws ecs update-service \
  --cluster startup-scout-cluster \
  --service startup-scout-service \
  --force-new-deployment
```

### Download the live Excel for a backup
```bash
# SSH into a bastion (or run an ECS Exec session) and copy from EFS
aws ecs execute-command \
  --cluster startup-scout-cluster \
  --task TASK_ID \
  --container startup-scout \
  --interactive \
  --command "/bin/sh"

# Inside the container:
cp /data/Startup\ database.xlsx /tmp/backup-$(date +%Y%m%d).xlsx
# Then use `aws s3 cp` to push to S3
```

### Scale up if the app is slow under load
```bash
aws ecs update-service \
  --cluster startup-scout-cluster \
  --service startup-scout-service \
  --desired-count 2
```
