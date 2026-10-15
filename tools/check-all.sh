#!/bin/sh
# Единая точка входа для локальной проверки — ровно те же шаги, что в CI
# (.github/workflows/vba-lint.yml). Запуск: ./tools/check-all.sh из корня или
# из любого каталога. Возврат: 0 — всё зелено, иначе код первой упавшей проверки.
set -e
cd "$(dirname "$0")/.."

echo "== 1/5 vba-lint: статика Module1.bas =="
python3 tools/vba-lint.py

echo
echo "== 2/5 parse-spec: разбор чисел из текста =="
python3 tools/parse-spec.py

echo
echo "== 3/5 plan-spec: укладчик реек =="
python3 tools/plan-spec.py

echo
echo "== 4/5 check-readme-map: карта процедур в README =="
python3 tools/check-readme-map.py

echo
echo "== 5/5 check-sync: зеркала, константы и обещания README =="
python3 tools/check-sync.py

echo
echo "все проверки пройдены"
