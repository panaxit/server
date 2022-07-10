<!--#include file="vbscript.asp"-->
<%
Dim RegEx: Set RegEx = New RegExp
With RegEx
    .Pattern = "(\[[^\[]*\])+"
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
End With
DIM content_type: content_type=Request.ServerVariables("HTTP_ACCEPT")

Dim objConn, strFile
Dim intCampaignRecipientID

strRFC = Request.QueryString("rfc")
strFullPath = REPLACE(Request.QueryString("full_path"),"/","\")
strFile = Right(strFullPath, InStr(StrReverse(strFullPath),"\")-1)

base_folder="D:\Dropbox (Personal)\Proyectos\Petro\FilesRepository"
    
IF INSTR(UCASE(strFullPath),UCASE("C:\fakepath"))>0 THEN
    strFullPath=base_folder & strFile
ELSE
    strFullPath=base_folder & "\SociosDeNegocios\Autorizado\"&strRFC & "\" & strFile
END IF

Dim objStream
If strFile <> "" Then
    Response.Buffer = False
    Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Type = 1 'adTypeBinary
    objStream.Open
    objStream.LoadFromFile(strFullPath)
    Response.ContentType = "application/x-unknown"
    Response.Addheader "Content-Disposition", "attachment; filename=" & strFile
    Response.BinaryWrite objStream.Read
    objStream.Close
    Set objStream = Nothing
End If
%>
