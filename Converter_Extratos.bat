@echo on
cd /d "%~dp0"

echo ============================================
echo   Conversor de Extrato PDF para Excel - Jean Vieira
echo ============================================

for %%f in (*.pdf) do (
    python converter.py "%%f"
)

echo ============================================
echo   Concluido. Os arquivos .xlsx foram gerados
echo   nesta mesma pasta.
echo ============================================
pause
