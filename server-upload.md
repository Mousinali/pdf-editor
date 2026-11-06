# Deploying to an Ubuntu server on Oracle Cloud (OCI)

> `server.py`'s `/convert` route now shells out to LibreOffice headless (`soffice`)
> instead of the old Windows-only `win32com`/Word COM automation, so it runs on
> Ubuntu. Make sure `libreoffice` is installed (step 2) — the route will fail
> without it. `/pdf2word` already used the pure-Python `pdf2docx` library and
> needs no changes.

## 1. Provision the instance
- OCI Console → Compute → Create Instance → Ubuntu 22.04/24.04 image.
- Add your SSH public key, note the public IP.
- OCI Console → Networking → VCN → Security Lists → add ingress rules for
  ports **22** (SSH), **80**/**443** (HTTP/HTTPS), and the app port if testing
  directly (e.g. 8000).

## 2. Connect and install base packages
```bash
ssh ubuntu@<INSTANCE_PUBLIC_IP>
sudo apt update && sudo apt upgrade -y
sudo apt install -y python3-venv python3-pip mysql-server nginx libreoffice
```
(`libreoffice` provides `soffice`, used by `/convert` for Word→PDF.)

## 3. Get the code from GitHub
On the server:
```bash
sudo apt install -y git
git clone <YOUR_GITHUB_REPO_URL> ~/pdf-app
cd ~/pdf-app
```
To deploy a later update:
```bash
cd ~/pdf-app
git pull
sudo systemctl restart pdf-app
```

## 4. Set up the app on the server
```bash
cd ~/pdf-app
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt gunicorn
```

Create `.env` from `.env.example` with real MySQL credentials:
```bash
cp .env.example .env
nano .env
```

Configure MySQL:
```bash
sudo mysql -e "CREATE DATABASE document_converter; \
  CREATE USER 'pdfapp'@'localhost' IDENTIFIED BY '<password>'; \
  GRANT ALL ON document_converter.* TO 'pdfapp'@'localhost';"
```

## 5. Run with gunicorn as a systemd service
`/etc/systemd/system/pdf-app.service`:
```ini
[Unit]
Description=Pdf Convertor
After=network.target mysql.service

[Service]
User=ubuntu
WorkingDirectory=/home/ubuntu/pdf-app
EnvironmentFile=/home/ubuntu/pdf-app/.env
ExecStart=/home/ubuntu/pdf-app/venv/bin/gunicorn -w 4 -b 127.0.0.1:8000 server:app
Restart=always

[Install]
WantedBy=multi-user.target
```
```bash
sudo systemctl daemon-reload
sudo systemctl enable --now pdf-app
sudo systemctl status pdf-app
```

## 6. Reverse proxy with nginx
`/etc/nginx/sites-available/pdf-app`:
```nginx
server {
    listen 80;
    server_name <INSTANCE_PUBLIC_IP_OR_DOMAIN>;

    client_max_body_size 50M;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```
```bash
sudo ln -s /etc/nginx/sites-available/pdf-app /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
```

## 7. (Optional) HTTPS
```bash
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d <your-domain>
```

## 8. Verify
```bash
curl http://<INSTANCE_PUBLIC_IP>/
```
Test each tool endpoint from the browser UI (`index.html`) served at `/`.
