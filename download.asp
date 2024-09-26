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

DIM authorization: authorization = Request.ServerVariables("HTTP_AUTHORIZATION")
If (authorization<>"") Then
    login
End if

Dim configuration: Set configuration = getConfiguration()

Server.ScriptTimeOut=1200
response.Buffer=true
IF NOT(Session("AccessGranted")) THEN 
    Response.ContentType = "application/json"
    Response.CharSet = "ISO-8859-1"
    Response.Status = "401 Unauthorized" %>
    {
        "status":"unauthorized"
      , "message":""
    }
    <% 
    response.end
END IF

Dim objConn, strFile
Dim intCampaignRecipientID

strFullPath = REPLACE(Request.QueryString("file"),"/","\")
if (InStr(StrReverse(strFullPath),"\")<>0) then
    strFile = Right(strFullPath, InStr(StrReverse(strFullPath),"\")-1)
else 
    strFile = strFullPath
end if

base_folder = configuration.getAttribute("root_folder")
IF isnull(base_folder) THEN
    strFullPath = server.MapPath(".") & strFile
ELSE
    strFullPath = base_folder & strFile
END IF

ON ERROR RESUME NEXT
DIM xmlDoc:	set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = false
xmlDoc.Load(strFullPath)
response.CodePage = 65001
response.CharSet = "UTF-8"
response.ContentType = "text/xml" 
'if (xmlDoc.selectSingleNode("//x:source/*") is nothing) then
    response.write xmlDoc.xml
'end if
'Dim objStream
'If strFile <> "" Then
'    Response.Buffer = False
'    Set objStream = Server.CreateObject("ADODB.Stream")
'    objStream.Type = 1 'adTypeBinary
'    objStream.Open
'    objStream.LoadFromFile(strFullPath)
'    Response.ContentType = "application/x-unknown"
'    Response.Addheader "Content-Disposition", "attachment; filename=" & strFile
'    Response.BinaryWrite objStream.Read
'    objStream.Close
'    Set objStream = Nothing
'End If
IF Err.Number<>0 THEN
    response.write "<!--<WARNING>"& Err.Description &"</WARNING>-->"
END IF
ON ERROR GOTO 0
%>
