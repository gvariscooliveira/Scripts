@echo on
set dnsserver=XX.XX.XX.XX
set dnsserver2=XX.XX.XX.XX
for /f "tokens=1,2,3*" %%i in ('netsh interface show interface') do (
 if %%i EQU Habilitado (
 netsh interface ipv4 set dnsserver name="%%l" static %dnsserver% both
 netsh interface ipv4 add dnsserver name="%%l" %dnsserver2% index=2
 )
)