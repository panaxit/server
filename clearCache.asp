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
Response.CharSet = "ISO-8859-1"
ON ERROR  RESUME NEXT
IF content_type<>"" THEN
    Response.ContentType = content_type
END IF
'IF INSTR(content_type,"xml") THEN
'Response.ContentType = "text/xml"
DIM xmlDoc
Set xmlDoc=Server.CreateObject("Microsoft.XMLDOM")
xmlDoc.async="false"
xmlDoc.load(Request)
DIM file_location
file_location=server.MapPath(".")&"\..\cache\"&request.querystring("file_name")&".xml"

dim fs
Set fs=Server.CreateObject("Scripting.FileSystemObject")
if fs.FileExists(file_location) then
  fs.DeleteFile(file_location)
end if
set fs=nothing

IF Err.Number<>0 THEN
    Response.Clear()
	ErrorDesc=Err.Description
    Response.Status = "409 Conflict"
    response.write "//"&file_location
    DIM message: message=RegEx.Replace(Err.Description, "")
    IF INSTR(content_type,"xml")>0 THEN %>
<x:message xmlns:x="http://panax.io/xover" type="exception"><%= message %></x:message>
<%  ELSEIF INSTR(content_type,"json")>0 THEN
    Response.ContentType = "application/json" %>
{"status":"exception","message":"<%= REPLACE(message, """", "\""") %>"}
<%  ELSE 
    Response.ContentType = "application/javascript" %>
    this.status='exception';
    this.message="<%= REPLACE(message, """", "\""") %>";
    <%
    END IF
    response.end
ELSE
    IF INSTR(content_type,"xml")>0 THEN %>
<x:message xmlns:x="http://panax.io/xover" type="success"><%= message %></x:message>
<%  ELSEIF INSTR(content_type,"json")>0 THEN
    Response.ContentType = "application/json" %>
{"status":"success"}
<%  ELSE 
    Response.ContentType = "application/javascript" %>
this.status='success'
<%  END IF
END IF %>