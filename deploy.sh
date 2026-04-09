#!/usr/bin/env bash
# =============================================================================
#  MEIL MDM — Ubuntu Server Deployment Script
#  Run as root or with sudo:  sudo bash deploy.sh
# =============================================================================
set -euo pipefail

APP_DIR="/opt/meil"
BACKEND_DIR="$APP_DIR/meil_backend"
FRONTEND_DIR="$APP_DIR/meil_frontend"
BACKEND_PORT=8000
FRONTEND_PORT=3000
APP_USER="meil"

GREEN='\033[0;32m'; YELLOW='\033[1;33m'; RED='\033[0;31m'; NC='\033[0m'
info()    { echo -e "${GREEN}[INFO]${NC} $*"; }
warn()    { echo -e "${YELLOW}[WARN]${NC} $*"; }
section() { echo -e "\n${GREEN}━━━ $* ━━━${NC}"; }

# ─── 1. System packages ───────────────────────────────────────────────────────
section "1/7  Installing system packages"
apt-get update -y
apt-get install -y \
    python3 python3-pip python3-venv \
    postgresql postgresql-contrib \
    nginx git curl build-essential

# Node.js 20 via NodeSource
if ! command -v node &>/dev/null; then
    info "Installing Node.js 20..."
    curl -fsSL https://deb.nodesource.com/setup_20.x | bash -
    apt-get install -y nodejs
fi

info "Node  : $(node --version)"
info "Python: $(python3 --version)"
info "npm   : $(npm --version)"

# ─── 2. App user & directory ──────────────────────────────────────────────────
section "2/7  Setting up app user and directory"
id "$APP_USER" &>/dev/null || useradd -r -m -s /bin/bash "$APP_USER"
mkdir -p "$APP_DIR"

if [ ! -d "$APP_DIR/.git" ]; then
    read -rp "Enter git repo SSH/HTTPS URL: " REPO_URL
    git clone "$REPO_URL" "$APP_DIR"
else
    info "Repo already cloned. Pulling latest..."
    git -C "$APP_DIR" pull
fi

chown -R "$APP_USER":"$APP_USER" "$APP_DIR"

# ─── 3. PostgreSQL ────────────────────────────────────────────────────────────
section "3/7  PostgreSQL setup"
systemctl enable --now postgresql

read -rp "DB name   [meil_mdm]:       " DB_NAME;   DB_NAME="${DB_NAME:-meil_mdm}"
read -rp "DB user   [meil_user]:      " DB_USER;   DB_USER="${DB_USER:-meil_user}"
read -rsp "DB password: "                           DB_PASS;  echo

# Use 'su - postgres' to avoid pg_hba md5 auth issues with sudo -u postgres
PG_CMD="su - postgres -c"

$PG_CMD "psql -tc \"SELECT 1 FROM pg_user WHERE usename='$DB_USER'\"" | grep -q 1 \
    || $PG_CMD "psql -c \"CREATE USER $DB_USER WITH PASSWORD '$DB_PASS';\""

$PG_CMD "psql -tc \"SELECT 1 FROM pg_database WHERE datname='$DB_NAME'\"" | grep -q 1 \
    || $PG_CMD "psql -c \"CREATE DATABASE $DB_NAME OWNER $DB_USER;\""

$PG_CMD "psql -c \"GRANT ALL PRIVILEGES ON DATABASE $DB_NAME TO $DB_USER;\""
info "Database '$DB_NAME' ready."

# ─── 4. Backend ───────────────────────────────────────────────────────────────
section "4/7  Backend setup"
cd "$BACKEND_DIR"

# Generate a strong SECRET_KEY
SECRET_KEY=$(python3 -c "from django.core.management.utils import get_random_secret_key; print(get_random_secret_key())" 2>/dev/null \
    || python3 -c "import secrets, string; print(''.join(secrets.choice(string.ascii_letters+string.digits+'!@#\$%^&*') for _ in range(50)))")

SERVER_IP=$(hostname -I | awk '{print $1}')

if [ ! -f "$BACKEND_DIR/.env" ]; then
    cat > "$BACKEND_DIR/.env" << EOF
SECRET_KEY=$SECRET_KEY
DEBUG=False
ALLOWED_HOSTS=$SERVER_IP,localhost,127.0.0.1
DATABASE_URL=postgresql://$DB_USER:$DB_PASS@localhost:5432/$DB_NAME
EMAIL_HOST=mail.meghaeng.com
EMAIL_PORT=587
EMAIL_USE_TLS=True
EMAIL_HOST_USER=mdgtadmin@meghaeng.com
EMAIL_HOST_PASSWORD=
FRONTEND_BASE_URL=http://$SERVER_IP:$FRONTEND_PORT
MSG91_AUTH_KEY=
MSG91_SENDER_ID=MSG91
MSG91_TEMPLATE_ID=
MSG91_ROUTE=4
EOF
    warn "Review/update email password in $BACKEND_DIR/.env"
fi

sudo -u "$APP_USER" python3 -m venv "$BACKEND_DIR/venv"
sudo -u "$APP_USER" "$BACKEND_DIR/venv/bin/pip" install --upgrade pip
sudo -u "$APP_USER" "$BACKEND_DIR/venv/bin/pip" install -r "$BACKEND_DIR/requirements.txt"

# Run Django setup
sudo -u "$APP_USER" bash -c "
    cd $BACKEND_DIR
    source venv/bin/activate
    python manage.py migrate --no-input
    python manage.py seed_default_users
    python manage.py collectstatic --no-input --clear
"

# ─── 5. Frontend ──────────────────────────────────────────────────────────────
section "5/7  Frontend setup"
cd "$FRONTEND_DIR"

if [ ! -f "$FRONTEND_DIR/.env.production" ]; then
    cat > "$FRONTEND_DIR/.env.production" << EOF
NEXT_PUBLIC_API_BASE_URL=http://$SERVER_IP:$BACKEND_PORT
EOF
fi

sudo -u "$APP_USER" npm install --prefix "$FRONTEND_DIR"
sudo -u "$APP_USER" npm run build --prefix "$FRONTEND_DIR"

# ─── 6. Systemd services ──────────────────────────────────────────────────────
section "6/7  Systemd services"

cat > /etc/systemd/system/meil-backend.service << EOF
[Unit]
Description=MEIL MDM Backend (Daphne ASGI)
After=network.target postgresql.service

[Service]
Type=simple
User=$APP_USER
WorkingDirectory=$BACKEND_DIR
EnvironmentFile=$BACKEND_DIR/.env
ExecStart=$BACKEND_DIR/venv/bin/daphne -b 0.0.0.0 -p $BACKEND_PORT core.asgi:application
Restart=always
RestartSec=5
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
EOF

cat > /etc/systemd/system/meil-frontend.service << EOF
[Unit]
Description=MEIL MDM Frontend (Next.js)
After=network.target meil-backend.service

[Service]
Type=simple
User=$APP_USER
WorkingDirectory=$FRONTEND_DIR
Environment=NODE_ENV=production
Environment=PORT=$FRONTEND_PORT
ExecStart=/usr/bin/node node_modules/.bin/next start -p $FRONTEND_PORT
Restart=always
RestartSec=5
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable meil-backend meil-frontend
systemctl restart meil-backend meil-frontend

# ─── 7. Nginx ─────────────────────────────────────────────────────────────────
section "7/7  Nginx configuration"

cat > /etc/nginx/sites-available/meil << EOF
server {
    listen 80;
    server_name $SERVER_IP _;

    # Frontend
    location / {
        proxy_pass http://127.0.0.1:$FRONTEND_PORT;
        proxy_http_version 1.1;
        proxy_set_header Upgrade \$http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host \$host;
        proxy_cache_bypass \$http_upgrade;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
    }

    # Backend API
    location /api/ {
        rewrite ^/api/(.*)\$ /\$1 break;
        proxy_pass http://127.0.0.1:$BACKEND_PORT;
        proxy_http_version 1.1;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
    }

    # WebSocket support for Django Channels
    location /ws/ {
        proxy_pass http://127.0.0.1:$BACKEND_PORT;
        proxy_http_version 1.1;
        proxy_set_header Upgrade \$http_upgrade;
        proxy_set_header Connection "Upgrade";
        proxy_set_header Host \$host;
    }

    client_max_body_size 50M;
}
EOF

ln -sf /etc/nginx/sites-available/meil /etc/nginx/sites-enabled/meil
rm -f /etc/nginx/sites-enabled/default
nginx -t && systemctl reload nginx

# ─── Done ─────────────────────────────────────────────────────────────────────
echo ""
echo -e "${GREEN}═══════════════════════════════════════════════════════════════${NC}"
echo -e "${GREEN}  MEIL MDM deployment complete!${NC}"
echo -e "${GREEN}═══════════════════════════════════════════════════════════════${NC}"
echo ""
echo "  Frontend  →  http://$SERVER_IP"
echo "  Backend   →  http://$SERVER_IP:$BACKEND_PORT"
echo ""
echo "  Default login credentials:"
echo "    SUPERADMIN : superadmin@meil.com  /  SuperAdmin@123"
echo "    ADMIN      : admin@meil.com       /  Admin@123"
echo "    MDGT       : mdgt@meil.com        /  Mdgt@123"
echo "    USER       : user@meil.com        /  User@123"
echo ""
echo "  Service logs:"
echo "    sudo journalctl -u meil-backend -f"
echo "    sudo journalctl -u meil-frontend -f"
echo ""
warn "Change default passwords after first login!"
