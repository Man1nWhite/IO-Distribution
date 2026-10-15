#!/bin/sh
# Обёртка для привычной команды ./tools/check-all.sh (из корня или из любого
# каталога). Шаги живут в tools/check-all.py — он кроссплатформенный (Windows:
# py tools\check-all.py), а его совпадение с CI сверяет tools/check-sync.py.
# Возврат: 0 — всё зелено, иначе код первой упавшей проверки.
exec python3 "$(dirname "$0")/check-all.py" "$@"
