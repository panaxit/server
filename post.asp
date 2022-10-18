<!--#include file="vbscript.asp"-->
<%
Dim RegEx: Set RegEx = New RegExp
With RegEx
    .Pattern = "'?(\.?\[[^\[]*\])+'?"
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
End With

function asyncCall(strUrl)
    Set xmlHttp = Server.Createobject("MSXML2.ServerXMLHTTP")
    xmlHttp.Open "GET", strUrl, False
    xmlHttp.setRequestHeader "User-Agent", "asp httprequest"
    xmlHttp.setRequestHeader "content-type", "application/x-www-form-urlencoded"
    xmlHttp.Send
    getHTML = xmlHttp.responseText
    xmlHttp.abort()
    set xmlHttp = Nothing   
end function 

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

'Functions
Sub manageError(Err)
    Response.CharSet = "UTF-8"
    Response.Clear()
    IF INSTR(Err.Description, "interbloqueo") THEN
        Response.Status = "423 Locked"
    ELSE 
        Response.Status = "409 Conflict"
    END IF

    DIM message: message=RegEx.Replace(Err.Description, "")
    IF message="" AND r>max_recordsets THEN
        message = "La solicitud devolvió más conjuntos de datos de los permitidos"
    ELSEIF INSTR(message,"SQL Server does not exist or access denied")>0 OR INSTR(message,"Communication link failure")>0 THEN
        Response.Status = "503 Service Unavailable" '"408 Request Timeout"
        AsyncCall "https://server.panax.io:8081/startSQL"
    ELSEIF INSTR(UCASE(message), UCASE("clave duplicada"))>0 THEN
		message="PRECAUCIÓN: No se puede insertar un registro duplicado."
	ELSEIF INSTR(UCASE(message), UCASE("La columna no admite valores NULL"))>0 THEN
		message="El campo no se puede quedar vacío"
	ELSE
		'message="El sistema no pudo completar el proceso y envió el siguiente mensaje: \n\n"&message
	END IF

    IF INSTR(Response.ContentType,"xml")>0 THEN 
        IF SESSION("user_login")="webmaster" OR SESSION("debug") THEN
            response.write "<!--"&strSQL&"-->"
        END IF
%>
<?xml-stylesheet type="text/xsl" href="message.xslt" role="message" target="body" action="append"?>
<x:message xmlns:x="http://panax.io/xover" x:id="message_<%= REPLACE(REPLACE(REPLACE(NOW(),":",""),"/","")," ","_") %>" type="exception"><%= REPLACE(REPLACE(message,">","&gt;"),"<","&lt;") %></x:message>
<%  ELSEIF INSTR(Response.ContentType,"json")>0 THEN %>
//<%= strSQL  %>
<%  ELSE 
        Response.ContentType = "application/javascript" %>
        this.status='exception';
        this.message="<%= REPLACE(message, """", "\""") %>";<%  
        IF 1=1 OR session("debug")=TRUE THEN %>
            this.source="<%= REPLACE(strSQL, """", "\""") %>";
    <%  END IF 
    END IF 
End Sub
'TERMINAN FUNCIONES

Response.CodePage = 65001
Response.CharSet = "UTF-8"
DIM accept: accept=Request.ServerVariables("HTTP_ACCEPT")
IF accept="*/*" OR accept="*/*, */*" THEN
    Response.ContentType = "text/xml"
ELSE
    Response.ContentType = accept
END IF
DIM max_recordsets: max_recordsets = TRIM(Request.ServerVariables("HTTP_X_MAX_RECORDSETS"))
IF max_recordsets="" THEN
    max_recordsets = 1
END IF

'for each x in Request.ServerVariables
'  response.write("<B>" & x & ":</b> " & Request.ServerVariables(x) & "<p />")
'next
DIM oConfiguration:	set oConfiguration = Server.CreateObject("MSXML2.DOMDocument"): 
oConfiguration.Async = false: 
oConfiguration.setProperty "SelectionLanguage", "XPath"
oConfiguration.Load(Server.MapPath("../../.config/system.config"))

IF NOT(Session("AccessGranted")) THEN 
    Response.ContentType = "application/javascript"
    Response.CharSet = "ISO-8859-1"
    Response.Status = "401 Unauthorized"%>
    this.status='unauthorized';
    this.message="<%= REPLACE(REPLACE(ErrorDesc, "[Microsoft][ODBC SQL Server Driver][SQL Server]", ""), """", "\""") %>";
    <%
    response.end
END IF
DIM StrCnn: StrCnn = "driver={SQL Server};server="&SESSION("secret_server_id")&";uid="&SESSION("secret_database_user")&";pwd="&SESSION("secret_database_password")&";database="&SESSION("secret_database_name")
Set oCn = Server.CreateObject("ADODB.Connection")
oCn.ConnectionTimeout = 0
oCn.CommandTimeout = 0
oCn.Open StrCnn
IF Err.Number<>0 THEN
    Response.Clear()
    Response.Status = "401 Unauthorized"
	ErrorDesc=Err.Description
    %>
    this.status='unauthorized';
    this.message="<%= REPLACE(REPLACE(ErrorDesc, "[Microsoft][ODBC SQL Server Driver][SQL Server]", ""), """", "\""") %>";<%
    response.end
END IF
oCn.Execute("SET LANGUAGE SPANISH")

DIM debug: debug=Request.ServerVariables("HTTP_X_DEBUGGING")
'Response.CodePage = 65001
'Response.CharSet = "UTF-8"
Response.CharSet = "ISO-8859-1"
ON ERROR RESUME NEXT

DIM xmlDoc
Set xmlDoc=Server.CreateObject("Microsoft.XMLDOM")
xmlDoc.async="false"
xmlDoc.load(request)
'xmlDoc.load(server.MapPath("..\panax\post from v12 entity - ejemplo Puesto.xml"))

IF ISNULL(session("user_id")) = FALSE THEN
    xmlDoc.selectSingleNode("//x:source/*").setAttribute "session:user_id", session("user_id")
END IF
DIM xmlSubmit
Set xmlSubmit=Server.CreateObject("Microsoft.XMLDOM")
xmlSubmit.async="false"
IF xmlDoc.selectSingleNode("//x:submit/*") IS NOTHING THEN
    DIM xmlSource
    Set xmlSource=Server.CreateObject("Microsoft.XMLDOM")
    xmlSource.async="false"
    xmlSource.loadXML(xmlDoc.selectSingleNode("//x:source/*").xml)

    Set xslDoc=Server.CreateObject("Microsoft.XMLDOM")
    xslDoc.async="false"
    xslDoc.load(server.MapPath("..\panax\post.v12.xslt"))

    xmlSubmit.loadXML("<x:submit xmlns:x=""http://panax.io/xover"">"&xmlSource.transformNode(xslDoc)&"</x:submit>")

    dim root_node:  set root_node = xmlDoc.selectSingleNode("//x:post")
    root_node.insertBefore xmlSubmit.firstChild, root_node.firstChild
ELSE
    xmlSubmit.loadXML(xmlDoc.selectSingleNode("//x:submit").xml)
END IF

DIM file_location, parent_folder
parent_folder=server.MapPath(".")&"\..\sessions\save\"
file_location=parent_folder&"user_"&session("user_id")&"_"&REPLACE(REPLACE(REPLACE(NOW(),":",""),"/","")," ","_")&".xml"
set fso=CreateObject("Scripting.FileSystemObject")
If  Not fso.FolderExists(parent_folder) Then      
  CreateFolder(parent_folder)
End If

'response.write server.MapPath(".")&"\custom\sessions\save\user_"&session("user_id")&"_"&REPLACE(REPLACE(REPLACE(NOW(),":",""),"/","")," ","_")&".xml"
xmlDoc.save file_location
IF xmlDoc.selectSingleNode("//x:submit/*/*") IS NOTHING THEN
    Response.Status = "304 Not Modified"
    Response.End
END IF

DIM strSQL
'strSQL="SET NOCOUNT ON; DECLARE @success BIT; EXEC [#panax].[RegisterTransaction] '"&REPLACE(REPLACE(URLDecode(xmlDoc.selectSingleNode("//x:submit/*").xml), "&", "&amp;"), "'", "''")&"', @source='"&REPLACE(REPLACE(URLDecode(xmlDoc.selectSingleNode("//x:source/*").xml), "&", "&amp;"), "'", "''")&"', @user_id="& session("user_id") &", @exec=1, @success=@success;-- SELECT @success"

'base_folder="D:\Dropbox (Personal)\Proyectos\Petro\FilesRepository"
DIM xRepository: 
base_folder=server.MapPath("../FilesRepository")
IF (NOT(oConfiguration.documentElement) IS NOTHING) THEN
    SET xRepository=oConfiguration.documentElement.selectSingleNode("/configuration/Repositories/Folder[@Id='upload']")
    IF NOT (xRepository IS NOTHING) THEN
        base_folder=xRepository.getAttribute("Location")
    END IF
END IF

user_id = session("user_id")
if user_id = "" or ISNULL(user_id) = TRUE then
    user_id = "-1"
end if
strSQL="SET NOCOUNT ON; DECLARE @success BIT, @transaction_id INT, @response XML; EXEC [#panax].[RegisterTransaction] '"&REPLACE(REPLACE(REPLACE(xmlDoc.selectSingleNode("//x:submit/*").xml, "&", "&amp;"), "'", "''"),"C:\fakepath",base_folder)&"', @source='"&REPLACE(REPLACE(xmlDoc.selectSingleNode("//x:source/*").xml, "&", "&amp;"), "'", "''")&"', @user_id="& user_id &", @exec=1, @success=@success OUTPUT, @transaction_id=@transaction_id OUTPUT, @response=@response OUTPUT; SELECT success=@success, transaction_id=@transaction_id, @response FOR XML PATH('response'), type"

strSQL="BEGIN TRY "&strSQL&" END TRY BEGIN CATCH DECLARE @Message NVARCHAR(MAX); SELECT @Message=ERROR_MESSAGE(); EXEC [$Table].[getCustomMessage] @Message=@Message, @Exec=1; END CATCH"

IF eval(debug) THEN %>
<!-- <%= strSQL  %> -->
<%  Response.Status = "449 Retry With Debugging Disabled"
    Response.End
END IF

'response.write(REPLACE(strSQL, "<", "&lt;")): response.end
'DIM recordset:	Set recordset = Server.CreateObject("ADODB.RecordSet")
'recordset.CursorLocation 	= 3
'recordset.CursorType 		= 3
set recordset=oCn.execute(strSQL)
IF SESSION("user_login")="webmaster" OR SESSION("debug") THEN
    IF INSTR(Response.ContentType,"xml")>0 THEN
        Response.CodePage = 65001
        Response.CharSet = "UTF-8"
        response.ContentType = "text/xml" 
        response.write "<!--"&strSQL&"-->"
    END IF
END IF
DIM r: r=0
DIM oField, sDataType, sValue 
DIM oXMLFile:	set oXMLFile = Server.CreateObject("Microsoft.XMLDOM")
oXMLFile.Async = false
DO
    r = r+1
    IF Err.Number<>0 OR r>max_recordsets THEN 
        manageError(Err)
    ELSEIF recordset.fields.Count>0 THEN 
        IF NOT (recordset.BOF and recordset.EOF) THEN 
            IF INSTR(Response.ContentType,"xml")>0 THEN
                    Response.CodePage = 65001
                    Response.CharSet = "UTF-8"
                    response.ContentType = "text/xml" 
                    'response.Write("<?xml version='1.0' encoding='UTF-8'?>")
                    oXMLFile.LoadXML(recordset(0))
                    IF oXMLFile.documentElement IS NOTHING THEN
                        IF Request.ServerVariables("HTTP_ROOT_NODE")<>"" THEN %>
    <<%= Request.ServerVariables("HTTP_ROOT_NODE") %> xmlns:x="http://panax.io/xover" xmlns:source="http://panax.io/fetch/request" />
    <%                  ELSE
                             Response.Status = "204 No Content"
                        END IF
                    ELSEIF NOT(oXMLFile.selectSingleNode("//result[@status='error']") IS NOTHING) THEN
                        Response.Status = "409 Conflict"
                    END IF
                    'oXMLFile.loadXML(oXMLFile.transformNode(xslValues))
                    xslFile=server.MapPath(".")&"\normalize_namespaces.xslt"
                    Set xslDoc=Server.CreateObject("Microsoft.XMLDOM")
                    xslDoc.async="false"
                    xslDoc.load(xslFile)
                    oXMLFile.loadXML(oXMLFile.transformNode(xslDoc))
                    'response.write "  Cache-Response: "&Request.ServerVariables("Cache-Response")&"-->"
                    'response.write "<!-- Cache-Response: "&Request.ServerVariables("HTTP_CACHE_RESPONSE")&"-->"
                    IF Request.ServerVariables("HTTP_CACHE_RESPONSE")="true" THEN
                        oXMLFile.save file_location
                        'response.write "<!-- Saved: "&file_location&"-->"
                    END IF
                    response.write oXMLFile.xml
            ELSE %>
                this.contentType = '<%= Response.ContentType %>'
	            this.status='success'
	            this.recordSet=new Array()
        <%	        DO UNTIL recordset.EOF %>
                var record = {}
            <%      FOR EACH oField IN recordset.fields 
                        IF oField.name="" THEN 
                        END IF 
                        IF TypeName(oField)="Field" THEN 
                            sDataType=TypeName(oField.value): sValue=oField.value 
                        ELSE 
                            sDataType=TypeName(oField): sValue=oField 
                        END IF %>
		                    record["<%= oField.name %>"]=<% SELECT CASE UCASE(sDataType): CASE "NULL": %>null<% CASE "BOOLEAN": %><% IF sValue THEN %>true<% ELSE %>false<% END IF %><% CASE ELSE %>"<%= RTRIM(REPLACE(replaceMatch(sValue, "["&chr(13)&""&chr(10)&""&vbcr&""&vbcrlf&"]", ""&vbcrlf),"""", """")) %>"<% END SELECT %>; 
	                    <% NEXT %>
                    this.recordSet.push(record)
                    <% recordset.MoveNext
 	                LOOP %>
    <% recordset.Close 
             END IF 
        ELSE 
            IF NOT(debug) THEN
                Response.Status = "204 No Content" 
            END IF
        END IF 
    ELSE 
        IF INSTR(Response.ContentType,"xml")>0 THEN
            Response.CodePage = 65001
            Response.CharSet = "UTF-8"
            response.ContentType = "text/xml" 
        %><?xml-stylesheet type="text/xsl" href="message.xslt" role="message" target="body" action="append"?>
<x:message xmlns:x="http://panax.io/xover" x:id="message_<%= REPLACE(REPLACE(REPLACE(NOW(),":",""),"/","")," ","_") %>" type="success">El proceso ha terminado</x:message>
        <% ELSE %>
	        this.status='success'
	        this.recordSet=new Array()<% 
        END IF
    END IF 
    set recordset = recordset.nextRecordSet
LOOP UNTIL recordset is nothing or r>max_recordsets
If oCn.State = 1 THEN
    oCn.Close
END IF
%>