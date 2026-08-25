# DOffice Auto — cai lai sach tren Windows 10 bang Command Prompt

Tai lieu nay danh cho may chi co CMD, khong can PowerShell.

## 0. Nguyen tac an toan

Bo cu chi go task va tien trinh DOffice, khong xoa ngay thu muc code cu. Sau
khi ban moi chay thanh cong, co the doi ten thu muc cu thanh `.backup` mot vai
ngay roi moi xoa neu chac chan khong can du lieu.

Khong dung lenh `taskkill /IM python.exe /F` vi co the tat ung dung Python khac.
Bo `uninstall_web_startup.cmd` chi tim PID dang nghe cong 8877 va chi dung neu
PID do la `python.exe` hoac `pythonw.exe`.

## 1. Tat va go bo ban cu

Mo Command Prompt. Chuyen vao thu muc ban moi vua giai nen, sau do chay:

```bat
uninstall_web_startup.cmd
```

Neu khong co file nay trong ban cu, go thu cong:

```bat
schtasks /End /TN "DOffice Web Dashboard"
schtasks /Delete /F /TN "DOffice Web Dashboard"
netstat -ano | findstr :8877
```

Neu lenh cuoi tra ve PID dang `LISTENING`, xem dung tien trinh:

```bat
tasklist /FI "PID eq 1234"
```

Chi khi do la `python.exe` hoac `pythonw.exe` cua DOffice moi dung:

```bat
taskkill /PID 1234 /F
```

Kiem tra lai den khi khong con ket qua:

```bat
netstat -ano | findstr :8877
```

## 2. Giai nen ban moi

Giai nen vao thu muc moi, vi du:

```text
D:\DO_Auto_v1.1.3_fix2
```

Khong giai nen de chong len thu muc cu. Khong copy `config.py` cu de de phong
dua loi cu tro lai. Neu can giu du lieu, chi copy rieng cac thu muc `data` va
file `playwright\.auth\state.json` sau khi ban moi da chay thu.

## 3. Cai Python va thu vien

Can Python 3.11 hoac 3.12. Khi cai Python, tick `Add Python to PATH`.

Trong thu muc ban moi, chay:

```bat
setup_windows_cmd.bat
```

Script se tao `.venv`, cai Flask/openpyxl/Playwright va Chromium.

## 4. Tao phien dang nhap

Chay:

```bat
login_doffice.bat
```

Dang nhap trong cua so Chromium, quay lai CMD va nhan Enter. Phien duoc luu
vao `playwright\.auth\state.json`.

## 5. Chay thu web bang CMD

Chay:

```bat
run_web_cmd.bat
```

Mo `http://127.0.0.1:8877`. Chay mot tac vu test an toan. Neu moi thu on, bam
Ctrl+C trong cua so CMD de dung server thu.

## 6. Cai web chay nen cung Windows

Chay:

```bat
install_web_startup.cmd
```

Neu bi tu choi quyen, mo Command Prompt bang `Run as administrator` va chay lai.
Task dung `pythonw.exe`, nen khong mo cua so CMD. Khi dang nhap Windows, web se
tu chay tai `http://127.0.0.1:8877`.

Kiem tra task va cong bang:

```bat
check_web_cmd.cmd
```

## 7. Neu lai gap loi

Chay `uninstall_web_startup.cmd`, roi chay `run_web_cmd.bat` de xem traceback
truc tiep trong CMD. Ban sua moi da xu ly truong hop `pythonw.exe` khong co
`stdout`, nen khong con loi `'NoneType' object has no attribute 'write'`.

## 8. Go bo hoan toan sau khi da xac nhan ban moi

1. Chay `uninstall_web_startup.cmd`.
2. Xac nhan `netstat -ano | findstr :8877` khong con PID.
3. Doi ten thu muc cu thanh `DO_Auto_old_backup` va giu lai vai ngay.
4. Khi chac chan khong can, xoa thu muc backup bang File Explorer.

## 9. Khoi phuc

Neu ban moi co van de, dung `uninstall_web_startup.cmd`, sau do mo lai thu muc
backup va chay `run_web_cmd.bat`. Du lieu khong bi xoa boi cac script trong bo
cai nay.
