@echo off
REM ============================================================================
REM  install_offline.bat
REM  Офлайн-установка окружения конвертера.
REM  ЗАПУСКАТЬ НА ЦЕЛЕВОЙ МАШИНЕ БЕЗ ИНТЕРНЕТА (Windows x64, Python 3.12).
REM
REM  Создаёт .venv и ставит все пакеты из offline_kit\wheels без сети.
REM  Перед запуском папка offline_kit\wheels должна быть уже заполнена
REM  (см. build_offline_kit.bat на машине с интернетом).
REM ============================================================================
setlocal
cd /d "%~dp0.."

if not exist "offline_kit\wheels" (
  echo !!! Нет папки offline_kit\wheels.
  echo     Сначала соберите комплект на машине с интернетом: tools\build_offline_kit.bat
  exit /b 1
)

where python >nul 2>&1
if errorlevel 1 (
  echo !!! Python не найден в PATH. Установите Python 3.12 x64 и повторите.
  exit /b 1
)

echo === Склейка крупных wheel из кусков (.partNNN -^> .whl) ===
python tools\reassemble_wheels.py
if errorlevel 1 (echo !!! Склейка wheel не удалась. & exit /b 1)

echo === Создание виртуального окружения .venv ===
if not exist ".venv" python -m venv .venv

echo === Установка пакетов из offline_kit\wheels (без интернета) ===
.venv\Scripts\python.exe -m pip install --no-index --find-links offline_kit\wheels -r requirements-lock.txt
if errorlevel 1 (
  echo.
  echo !!! Установка не удалась. Проверьте, что offline_kit\wheels собран на
  echo     такой же платформе (Windows x64, Python 3.12).
  exit /b 1
)

echo.
echo === Готово! Запуск приложения: ===
echo     .venv\Scripts\python.exe main.py
echo Затем откройте в браузере http://127.0.0.1:8000
endlocal
