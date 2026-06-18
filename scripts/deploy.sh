#!/bin/bash
set -e

APP_DIR="/home/ec2-user/legal-app"
LOG_DIR="/home/ec2-user/logs"

echo "[deploy] Starting LexAI deployment..."

# Create log dir
mkdir -p "$LOG_DIR"

# Enter app directory
cd "$APP_DIR"

# Pull latest code
echo "[deploy] Pulling latest code..."
git pull origin main 2>&1 || echo "[deploy] Git pull failed (might be first deploy)"

# Install/update Python dependencies
echo "[deploy] Installing dependencies..."
pip3 install -r requirements.txt 2>&1 | tail -5

# Ensure .env exists
if [ ! -f .env ]; then
    echo "[deploy] Creating .env..."
    cat > .env << 'EOF'
AWS_REGION=ap-south-1
LEXAI_MODEL=apac.amazon.nova-pro-v1:0
FLASK_ENV=production
EOF
fi

# Restart the service
echo "[deploy] Restarting lexai service..."
sudo systemctl daemon-reload
sudo systemctl restart lexai

# Wait and check
sleep 3
if sudo systemctl is-active --quiet lexai; then
    echo "[deploy] Service is running."
    sudo systemctl status lexai --no-pager | head -5
else
    echo "[deploy] Service failed to start. Checking logs..."
    sudo journalctl -u lexai -n 20 --no-pager
    exit 1
fi

# Health check
echo "[deploy] Running health check..."
curl -sf http://localhost:5000/ && echo "[deploy] Health check passed!" || echo "[deploy] Health check failed (app may still be starting)"

echo "[deploy] Deployment complete!"
