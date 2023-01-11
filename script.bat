@if (@CodeSection == @Batch) @then
@echo off

set "ip=" & ::IP address of stor2rrd server
set "user=" & ::Username for stor2rrd
set "pass=" & ::Password for stor2rrd


set url="http://%ip%/stor2rrd-cgi/detail-graph.sh?host=ALL&type=ALL&name=cap_total&item=sum&time=d::::"

curl --user "%user%:%pass%" -o "Capacity.csv" %url%
echo "Prepared mail for send"

for /f "tokens=*" %%A in ('cscript //nologo dcript.vbs') do set "mail=%%A"

exit /b 0

goto :EOF
@end

// JScript section
var WshShell = WScript.CreateObject("WScript.Shell");
WshShell.SendKeys(WScript.Arguments(0));