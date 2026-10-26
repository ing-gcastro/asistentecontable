---
description: Commitea y publica cambios en main
---

1. Ejecutar `VERIFICAR.bat --sin-pausa`. Si no termina en `TESTS OK`, NO continuar: informar el error.
2. `git add -A`
3. `git commit -m "<tipo>: <resumen corto en español>"`
4. `git pull --rebase origin main`
5. `git push origin main`
6. Informar el hash del commit.
