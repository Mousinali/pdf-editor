# Deploying to an Ubuntu server on Oracle Cloud (OCI)

> **Blocker:** `server.py` uses `pywin32`/`win32com.client` (MS Word COM automation)
> for `/convert` and `/pdf2word`. That's Windows-only — it will not import or run
> on Ubuntu. Either drop those two routes or replace them with a Linux-compatible
> converter (e.g. LibreOffice headless: `soffice --headless --convert-to pdf`)
> before relying on them in production. Everything else (merge/split/compress/
> protect/unlock/rotate/watermark/editor) is pure Python and fine on Linux.

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
(`libreoffice` only needed if you replace the Word COM conversions.)

## 3. Upload the code
From your local machine:
```bash
scp -r /Users/mousinturain/Downloads/pdf ubuntu@<INSTANCE_PUBLIC_IP>:~/pdf-app
```
(or `git clone` if you push this to a repo instead.)

## 4. Set up the app on the server
```bash
cd ~/pdf-app
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt gunicorn
# remove pywin32 from requirements.txt first — it won't install on Linux
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
