echo off
mktyplib /nocpp /tlb CorType.tlb CorType.odl
if not errorlevel 1 goto end
pause
:end
regtlib CorType.tlb