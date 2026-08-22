# ============================================================
# Установка окружения курса «Практикум по ML и нейронным сетям»
# Требования: VS Code, Git, Miniconda уже установлены.
# Запуск: powershell -ExecutionPolicy Bypass -File .\install.ps1
# ============================================================

$ErrorActionPreference = "Stop"

# --- 0. Переменные (поменяйте при необходимости) ---
$EnvName    = "course_ml"                  # имя conda-окружения
$EnvPath    = $null                        # или путь, напр. "D:\envs\course_ml", если диск C: забит
$PipCache   = $null                        # или "D:\pip_cache", чтобы не забивать C:
# Пример для забитого C::
#   $EnvPath  = "D:\envs\course_ml"
#   $PipCache = "D:\pip_cache"
#   $env:TMP  = "D:\tmp"; $env:TEMP = "D:\tmp"

Write-Host "=== 1. Расширения VS Code ===" -ForegroundColor Cyan
$extensions = @(
    "ms-python.python",                    # Python
    "ms-python.vscode-pylance",            # Pylance (автодополнение, типы)
    "ms-toolsai.jupyter",                  # Jupyter (ноутбуки курса)
    "ms-python.vscode-python-manager",     # управление окружениями
    "charliermarsh.ruff",                  # линтер/форматтер
    "GrapeCity.gc-excelviewer"             # просмотр CSV
)
foreach ($ext in $extensions) {
    code --install-extension $ext
}

Write-Host "=== 2. Окружение conda (Python 3.12) ===" -ForegroundColor Cyan
if ($EnvPath) {
    conda create -p $EnvPath python=3.12 -y
    $python = Join-Path $EnvPath "python.exe"
} else {
    conda create -n $EnvName python=3.12 -y
    $python = (conda info --base).Trim() + "\envs\$EnvName\python.exe"
}

Write-Host "Использую Python: $python"
& $python --version
& $python -m pip install --upgrade pip

Write-Host "=== 3. Зависимости курса (requirements.txt) ===" -ForegroundColor Cyan
$pipArgs = @("-m", "pip", "install", "-r", "requirements.txt", "--no-input", "--progress-bar", "off")
if ($PipCache) { $pipArgs += @("--cache-dir", $PipCache) }
& $python @pipArgs

Write-Host "=== 4. Kaggle CLI (соревнования) ===" -ForegroundColor Cyan
& $python -m pip install kaggle --no-input --progress-bar off

Write-Host "=== 5. Проверка конфликтов ===" -ForegroundColor Cyan
& $python -m pip check

Write-Host "=== 6. Смоук-тесты всех модулей ===" -ForegroundColor Cyan
& $python tests\smoke_test.py

Write-Host ""
Write-Host "Готово. Ожидаемый результат выше: 'Пройдено: 9/9'." -ForegroundColor Green
Write-Host "Дальнейшие шаги:" -ForegroundColor Green
Write-Host "  - Kaggle API-токен: kaggle.json -> ~\.kaggle\ (kaggle.com -> Account -> Create New API Token)"
Write-Host "  - В VS Code: Ctrl+Shift+P -> Python: Select Interpreter -> $python"
Write-Host "  - Label Studio (М6): активировать окружение и запустить 'label-studio' (http://localhost:8080)"
