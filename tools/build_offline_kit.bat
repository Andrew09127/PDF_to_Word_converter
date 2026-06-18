@echo off
REM ============================================================================
REM  build_offline_kit.bat
REM  Сборка офлайн-комплекта Python-пакетов.
REM  ЗАПУСКАТЬ НА МАШИНЕ С ИНТЕРНЕТОМ (Windows x64, Python 3.12).
REM
REM  Скачивает все зависимости из requirements-lock.txt в offline_kit\wheels,
REM  включая torch+cpu из PyTorch CPU-индекса. Потом всю папку проекта
REM  (с offline_kit\, models\, frontend\dist\) переносят на офлайн-машину.
REM ============================================================================
setlocal
cd /d "%~dp0.."

set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo === Проверка Python ===
"%PY%" --version
echo.

if not exist "offline_kit\wheels" mkdir "offline_kit\wheels"

echo === Сборка wheels (это долго — torch и др. весят сотни МБ) ===
REM pip wheel (а не download): пакеты, у которых есть только исходник (напр.
REM   antlr4-python3-runtime==4.9.3, чистый Python), здесь же превращаются в wheel,
REM   чтобы на офлайн-машине ничего не собирать.
REM --only-binary для скомпилированных (llama/torch/torchvision): их строить из
REM   исходника нельзя (нужен компилятор; sdist llama.cpp ещё и ломает длину пути).
REM Доп. индексы: torch+cpu (PyTorch) и бинарный llama-cpp-python (abetlen).
"%PY%" -m pip wheel -r requirements-lock.txt -w offline_kit\wheels ^
  --only-binary=llama-cpp-python,torch,torchvision ^
  --extra-index-url https://download.pytorch.org/whl/cpu ^
  --extra-index-url https://abetlen.github.io/llama-cpp-python/whl/cpu
if errorlevel 1 (
  echo.
  echo !!! Скачивание не удалось. Проверьте интернет и что Python = 3.12 x64.
  exit /b 1
)

echo.
echo === Разбиение крупных wheel (^>95 МБ) на куски для лимита GitHub ===
"%PY%" tools\split_wheels.py
if errorlevel 1 (echo !!! Разбиение не удалось. & exit /b 1)

echo.
echo === Готово. offline_kit\wheels заполнен (крупные файлы — кусками .partNNN). ===
echo Закоммитьте offline_kit\ в git и запушьте. На офлайн-машине после git clone:
echo     tools\install_offline.bat
endlocal
