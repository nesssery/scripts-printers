# Utiliser une image Windows avec Python
FROM mcr.microsoft.com/windows/servercore:ltsc2019

# Installer Python
RUN powershell -Command \
    Invoke-WebRequest -Uri https://www.python.org/ftp/python/3.9.13/python-3.9.13-amd64.exe -OutFile python-installer.exe; \
    Start-Process -Wait -FilePath python-installer.exe -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1'; \
    Remove-Item python-installer.exe

# Installer pip et pywin32
RUN powershell -Command \
    Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -OutFile get-pip.py; \
    python get-pip.py; \
    pip install pywin32

# Définir le répertoire de travail
WORKDIR /app

# Copier le script et les dépendances
COPY print_watcher_win.py .
COPY requirements.txt .

# Installer les dépendances
RUN pip install --no-cache-dir -r requirements.txt

# Créer le dossier de logs
RUN mkdir \app\logs && type nul > \app\logs\output.log && type nul > \app\logs\error.log

# Lancer le script Python
CMD ["python", "print_watcher_win.py"]