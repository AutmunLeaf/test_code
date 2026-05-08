FROM python:3.11-slim-bullseye

# Предотвращает создание .pyc-файлов и буферизацию вывода
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Устанавливаем системные зависимости (включая .NET для Aspose.Cells)
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    wget \
    ca-certificates \
    fontconfig \
    libgdiplus \
    && rm -rf /var/lib/apt/lists/*

# Устанавливаем .NET Runtime 6.0 (LTS), необходимый для Aspose.Cells for Python via .NET
RUN wget https://packages.microsoft.com/config/debian/11/packages-microsoft-prod.deb -O packages-microsoft-prod.deb && \
    dpkg -i packages-microsoft-prod.deb && \
    rm packages-microsoft-prod.deb && \
    apt-get update && \
    apt-get install -y --no-install-recommends dotnet-runtime-6.0 && \
    rm -rf /var/lib/apt/lists/*

# Рабочая директория
WORKDIR /app

# Копируем зависимости
COPY requirements.txt .
# Копируем наши кастомные шрифты в стандартную директорию Linux
COPY ./fonts /usr/local/share/fonts

# Обновляем кэш шрифтов, чтобы система их "увидела"
RUN fc-cache -fv
RUN pip install --upgrade pip && pip install -r requirements.txt

# Копируем проект
COPY . .

# Создаём каталоги для статики и медиа, если они отсутствуют
RUN mkdir -p staticfiles media

# Команда по умолчанию (будет переопределена в compose)
CMD ["python", "manage.py", "runserver", "0.0.0.0:8000"]