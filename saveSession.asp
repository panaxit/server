<!--#include file="vbscript.asp"-->
<%
Function BytesToStr(bytes)
    Dim Stream
    Set Stream = Server.CreateObject("Adodb.Stream")
        Stream.Type = 1 'adTypeBinary
        Stream.Open
        Stream.Write bytes
        Stream.Position = 0
        Stream.Type = 2 'adTypeText
        Stream.Charset = "iso-8859-1"
        BytesToStr = Stream.ReadText
        Stream.Close
    Set Stream = Nothing
End Function

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
DIM payload
If Request.TotalBytes > 0 Then
    IF INSTR(Request.ServerVariables("HTTP_CONTENT_TYPE"),"xml") THEN
        Set payload=Server.CreateObject("Microsoft.XMLDOM")
        payload.async="false"
        payload.load(request)
    ELSE
        payload=BytesToStr(Request.BinaryRead(Request.TotalBytes))
    END IF
End If
DIM file_name: file_name = Request.ServerVariables("HTTP_X_FILE_NAME")
IF file_name="" THEN
    file_name = "user_"&session("user_id")
END IF
DIM file_location, parent_folder
parent_folder=server.MapPath(".")&"\..\..\sessions\"
file_location=parent_folder&file_name&".xml"
set fso=CreateObject("Scripting.FileSystemObject")
If  Not fso.FolderExists(parent_folder) Then      
  fso.CreateFolder (parent_folder)   
End If

IF INSTR(Request.ServerVariables("HTTP_CONTENT_TYPE"),"xml") THEN
    payload.save file_location
ELSE
    dim fs,tfile
    set fs=Server.CreateObject("Scripting.FileSystemObject")
    set tfile=fs.CreateTextFile(file_location)
    tfile.WriteLine(payload)
    tfile.close
    set tfile=nothing
    set fs=nothing
END IF

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