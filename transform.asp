<!--#include file="../../server/vbscript.asp"-->
<%  
Response.Buffer = True
DIM content_type: content_type=Request.ServerVariables("HTTP_CONTENT_TYPE")
'ON ERROR  RESUME NEXT

Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.ContentType = "text/html"
DIM lang: lang=request.QueryString("lang")
DIM xmlFile, xmlDoc
DIM Debug: Debug=CBOOL(request.QueryString("debug"))

DIM language
'IF lang="es" THEN 
'    language="spanish"
'ELSEIF lang="en" THEN 
'    language="english"
'ELSE
'    Response.Redirect "home.asp?lang=en"
'END IF
'xmlFile=server.MapPath(".")&"\..\resources\data\"&language&".xml"
'IF debug THEN
'    xmlDoc.selectSingleNode("/*").setAttribute "session:debug", "true"
'END IF
'IF request.servervariables("SERVER_NAME")="localhost" THEN
'    session("user_id") = "uriel_online@hotmail.com"
'    xmlDoc.selectSingleNode("/*").setAttribute "session:user_id", "uriel_online@hotmail.com"
'    xmlDoc.selectSingleNode("/*").setAttribute "session:user_name", "Uriel Gómez"
'    xmlDoc.selectSingleNode("/*").setAttribute "session:status", "authorized"
'    xmlDoc.selectSingleNode("/*").setAttribute "session:connected_to", "facebook"
'ELSE
'    xmlDoc.selectSingleNode("/*").setAttribute "session:thumbnail", URLDecode(session("thumbnail"))
'    xmlDoc.selectSingleNode("/*").setAttribute "session:user_id", session("user_id")
'    xmlDoc.selectSingleNode("/*").setAttribute "session:user_name", session("user_name")
'    xmlDoc.selectSingleNode("/*").setAttribute "session:status", "authorized"
'    xmlDoc.selectSingleNode("/*").setAttribute "session:connected_to", session("connected_to")
'END IF
'xmlDoc.loadXML(xmlDoc.xml)
'If (xmlDoc.parseError.errorCode <> 0) Then  
'   Set myErr = xmlDoc.parseError  
'   response.write myErr.reason
'   response.end
'End If  

'IF session("user_id")<>"" THEN
'    DIM reservation
'    xslFile=server.MapPath(".")&"\reservations\"&session("user_id")&".xml"
'    DIM fs
'    SET fs=Server.CreateObject("Scripting.FileSystemObject")
'    IF fs.FileExists(xslFile) then
'        DIM xmlData
'        Set xmlData=Server.CreateObject("Microsoft.XMLDOM")
'        xmlData.async="false"
'        'xmlData.setProperty("SelectionNamespaces", "xmlns:xsl='http://www.w3.org/1999/XSL/Transform'");
'        xmlData.load(xslFile)
'        Response.ContentType = "text/xml"
'        dim root_node:  set root_node = xmlDoc.selectSingleNode("wedding/*")
'        root_node.insertBefore xmlData.firstChild, root_node.firstChild
'        'response.Write xmlDoc.xml
'        'response.end
'    END IF
'    SET fs=NOTHING
'END IF
DIM myErr  
DIM xslFile, xslDoc
xslFile=server.MapPath(".")&"\xaml2html.xslt"
Set xslDoc=Server.CreateObject("Microsoft.XMLDOM")
xslDoc.async="false"
xslDoc.load(xslFile)

Set xmlDoc=Server.CreateObject("Microsoft.XMLDOM")
xmlDoc.async="false"
xmlDoc.SetProperty "SelectionLanguage", "XPath"
'xmlDoc.setProperty("SelectionNamespaces", "xmlns:xsl='http://www.w3.org/1999/XSL/Transform'");

xmlDoc.load(server.MapPath(".")&"\App.xaml")
xmlDoc.loadXML(xmlDoc.transformNode(xslDoc))
response.write xmlDoc.xml

xmlDoc.load(server.MapPath(".")&"\MainPage.xaml")
xmlDoc.loadXML(xmlDoc.transformNode(xslDoc))
response.write xmlDoc.xml

If (xmlDoc.parseError.errorCode <> 0) Then  
   Set myErr = xmlDoc.parseError  
   response.write myErr.reason
   response.end
End If  
%>