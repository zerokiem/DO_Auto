# DOffice Auto - image chay tren NAS Synology DS423 (ARM64) qua Container Manager.
#
# Dua tren image chinh thuc cua Playwright (da co san Chromium + moi thu vien he
# thong, ho tro ca linux/arm64) nen KHONG phai vat lon cai Chromium tren DSM.
# Code KHONG duoc COPY vao image ma se mount tu NAS luc chay (xem
# docker-compose.yml) -> sua code tren NAS khong can build lai image.
#
# KHONG dung apt-get: mang tu NAS ra ports.ubuntu.com hay bi BuildKit tren may
# nay timeout (network namespace rieng cua BuildKit, "docker run" thuong thi
# khong loi). May man la khong can apt-get chut nao: du lieu tzdata (zoneinfo)
# da co san trong image goc, va lich chay tu dong dung 1 thread Python ngay
# trong tien trinh web (do_auto/inprocess_scheduler.py) thay vi cron, nen cung
# khong can cai goi cron.
FROM mcr.microsoft.com/playwright/python:v1.55.0-noble

ENV TZ=Asia/Ho_Chi_Minh \
    DOFFICE_DATA_DIR=/data \
    PYTHONUNBUFFERED=1 \
    PYTHONUTF8=1

# zoneinfo da co san trong image goc (khong can apt-get install tzdata).
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime

# Image goc CO SAN browser Chromium (tai /ms-playwright, xem bien moi truong
# PLAYWRIGHT_BROWSERS_PATH) nhung KHONG co san goi pip "playwright" - phai tu
# cai. Ghim dung version 1.55.0 (khop voi tag image "v1.55.0-noble") de dung
# thang browser co san, KHONG tai lai browser moi (khong can "playwright
# install" vi browser da nam san trong image).
RUN pip install --no-cache-dir "flask>=3.0" "openpyxl>=3.1" "playwright==1.55.0"

WORKDIR /app

EXPOSE 8877

# Goi qua bash de khong phu thuoc bit thuc thi (+x) cua file mount tu NAS.
ENTRYPOINT ["bash", "/app/docker/entrypoint.sh"]
