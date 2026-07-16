#!/usr/bin/env bash
# Entrypoint cho container DOffice Auto: chay web dashboard o tien canh (PID 1).
# Lich chay tu dong (tab "Lich chay") dung 1 thread nen ngay trong tien trinh
# nay (xem do_auto/inprocess_scheduler.py) - KHONG can cron/goi ngoai nao.
set -e

cd /app

echo "DOffice web dashboard: http://0.0.0.0:8877"
exec python run_web.py
