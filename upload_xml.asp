<!--#include file="vbscript.asp"-->
<%
DIM content_type: content_type=Request.ServerVariables("HTTP_ACCEPT")
Response.CharSet = "ISO-8859-1"
'ON ERROR  RESUME NEXT'
'IF INSTR(content_type,"xml") THEN
'Response.ContentType = "text/xml"
DIM file_location, file_name
file_name = Request.ServerVariables("HTTP_FILE_NAME")
parent_folder = Request.ServerVariables("HTTP_PARENT_FOLDER")
DIM xmlDoc
Set xmlDoc=Server.CreateObject("Microsoft.XMLDOM")
xmlDoc.async="false"
xmlDoc.load(Request)
xmlDoc.setProperty "SelectionLanguage", "XPath"
xmlDoc.setProperty "SelectionNamespaces", "xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"""
Randomize

DIM ext
IF (TypeName(xmlDoc.selectSingleNode("/xsl:*"))<>"Nothing") THEN
    ext="xsl"
ELSEIF (TypeName(xmlDoc.selectSingleNode("/*[namespace-uri()='']"))<>"Nothing") THEN
    ext="xml"
ELSE
    ext="xml"
END IF

user_file_name="user_"&session("user_id")&"_"&REPLACE(REPLACE(REPLACE(NOW(),":",""),"/","")," ","_")&"_"& Rnd &"."&ext    
IF (file_name="") then
    file_name = user_file_name
END IF
IF parent_folder<>"" THEN
    parent_folder = server.MapPath(".")&"\..\"&parent_folder&"\"
ELSE
    parent_folder = server.MapPath(".")&"\..\"
END IF
file_location=parent_folder&file_name
set fso=CreateObject("Scripting.FileSystemObject")
If  Not fso.FolderExists(parent_folder) Then
    BuildFullPath parent_folder
  'fso.CreateFolder (parent_folder)   
End If
xmlDoc.save file_location 
IF NOT(file_name = user_file_name) THEN
    user_file_name="user_"&session("user_id")&"_"&REPLACE(REPLACE(REPLACE(NOW(),":",""),"/","")," ","_")&"_"& Rnd &"."&ext    
    file_location=server.MapPath(".")&"\..\sessions\save\"&user_file_name
    xmlDoc.save file_location
END IF
%>
this.fileName="<%= file_location %>";
<%
IF Err.Number<>0 THEN
	ErrorDesc=Err.Description
    %>
    this.status='exception';
    this.message="<%= REPLACE(REPLACE(ErrorDesc, "[Microsoft][ODBC SQL Server Driver][SQL Server]", ""), """", "\""") %>";
    <%
    response.end
ELSE
%>
	this.status='success'
<%
END IF %>