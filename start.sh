#!/bin/bash
set -e

export HOME=/tmp
mkdir -p /tmp/lo_profile

echo "啟動 LibreOffice headless UNO listener..."
soffice --headless --invisible --nocrashreport --nodefault --norestore \
  --nologo --nofirststartwizard \
  -env:UserInstallation=file:///tmp/lo_profile \
  --accept="socket,host=localhost,port=2002;urp;" &

# 等待 UNO socket 就緒
python3 - <<'PYEOF'
import socket, time
for _ in range(60):
    try:
        s = socket.create_connection(("localhost", 2002), timeout=1)
        s.close()
        print("LibreOffice UNO listener 已就緒")
        break
    except OSError:
        time.sleep(1)
else:
    raise SystemExit("等待 LibreOffice UNO listener 逾時")
PYEOF

echo "啟動 API 服務..."
exec gunicorn --bind 0.0.0.0:${PORT:-8080} --workers 2 --timeout 120 app:app
