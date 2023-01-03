'Convert csv file to html.

workspace = "" 'Change this path for your workspace
csvPath = workspace & "\Capacity.csv"
htmlPath = workspace & "\Capacity.html"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(csvPath, 1)
Set objHTML = objFSO.CreateTextFile(htmlPath, True)
Set objStoreIP = objFSO.OpenTextFile(workspace & "\storeIP.txt", 1)

storeIP = Split(objStoreIP.ReadAll, ";")

objHTML.WriteLine "<html><body><table border=1 style=""border-collapse: collapse; width: 100%;"">"
objHTML.WriteLine "<thead><tr>"

headers = Split(objFile.ReadLine, ";")
For Each header In headers
    objHTML.WriteLine "<th>" & header & "</th>"
Next

objHTML.WriteLine "<th>% Usado</th>"
objHTML.WriteLine "<th>% Libre</th>"
objHTML.WriteLine "<th>Store IP</th>"
objHTML.WriteLine "</tr></thead><tbody style=""text-align:center;"">"


do until objFile.AtEndOfStream

    objHTML.WriteLine "<tr>"
    body = Split(objFile.ReadLine, ";")
    count = 0
    For Each data In body
        if count = 4 then
            freeTB = CDbl(data)
            if freeTB <= 5 then
                objHTML.WriteLine "<td style=""background-color: #FF0000; color:white;"">" & data & "</td>"
            else
                objHTML.WriteLine "<td>" & data & "</td>"
            end if
            count = 0
        else
            objHTML.WriteLine "<td>" & data & "</td>"
        end if
        count = count + 1
    next

    totalDouble = CDbl(Replace(body(2),".",","))
    usedDouble = CDbl(Replace(body(3),".",","))
    freeDouble = CDbl(Replace(body(4),".",","))

    porcentUsed = Round((usedDouble / totalDouble) * 100,3)
    porcentFree = Round((freeDouble / totalDouble) * 100,3)

    objHTML.WriteLine "<td>" & porcentUsed & " %</td>"
    objHTML.WriteLine "<td>" & porcentFree & " %</td>"

    condition = searchInArray(storeIP, body(0))
    If condition(0) = 0 Then
        objHTML.WriteLine "<td style=""background-color: #FF0000;"">" & "NO ESPECIFICADO" & "</td>"
    Else
        objHTML.WriteLine "<td style=""background-color: #000000; color:white;""><a href=""http://" & Split(storeIP(condition(1)),"/")(1) & """ style=""text-decoration:none; color: white;"">" & Split(storeIP(condition(1)),"/")(1) & "</a></td>"
    End If

    objHTML.WriteLine "</tr>"

loop

objHTML.WriteLine "</tbody></table></body></html>"
objHTML.Close

'Prepare for send Mail

Set objMail = CreateObject("CDO.Message")

email_smtp = "" 'Change this for your smtp login email
password_smtp = "" 'Change this for your smtp login password

from = "" 'Change this for your email
destination = "" 'Change this for email destination

with objMail
    .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-mail.outlook.com"
    .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = email_smtp
    .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = password_smtp
    .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
    .Configuration.Fields.Update
end with

set htmlFile = objFSO.OpenTextFile(htmlPath, 1)

objMail.From = from
objMail.CC = destination
objMail.To = destination
objMail.Subject = "STORE ONCE REPORTE DE CAPACIDAD"
objMail.HtmlBody = htmlFile.ReadAll
objMail.AddAttachment htmlPath
objMail.Send


function searchInArray(arr,str)
    Dim Array
    Redim Array(2)
    for i = 0 to ubound(arr)
        if InStr(arr(i),str) > 0 then
            Array(0) = 1
            Array(1)= i
            searchInArray = Array
            exit function
        end if
    next
    Array(0) = 0
    Array(1) = 0
    searchInArray = Array
end function