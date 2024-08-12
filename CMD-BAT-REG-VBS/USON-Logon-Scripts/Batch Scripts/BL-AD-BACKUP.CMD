@ECHO OFF
for /f "skip=4 tokens=2 delims=:" %%g in ('"manage-bde -protectors -get c:"') do (
  manage-bde -protectors -adbackup C: -id%%g
  exit
  )