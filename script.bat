@if (@CodeSection == @Batch) @then
@echo off

set "ip=" & ::IP address of stor2rrd server
set "user=" & ::Username for stor2rrd
set "pass=" & ::Password for stor2rrd


set URL_1="http://%ip%/stor2rrd-cgi/detail-graph.sh?host=ALL&type=ALL&name=cap_total&item=sum&time=d::::"
set URL_2="http://%ip%/stor2rrd-cgi/detail-graph.sh?host=ALL&type=ALL&name=cap_pool&item=sum&time=d::::"

curl --user "%user%:%pass%" -o "Capacity_1.csv" %URL_1%
curl --user "%user%:%pass%" -o "Capacity_2.csv" %URL_2%


echo "Prepared mail for send"

for /f "tokens=*" %%A in ('cscript //nologo dcript.vbs') do set "mail=%%A"

exit /b 0

goto :EOF
@end

// JScript section
var WshShell = WScript.CreateObject("WScript.Shell");
WshShell.SendKeys(WScript.Arguments(0));