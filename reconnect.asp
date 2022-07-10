<% 
function asyncCall(strUrl)
    Set xmlHttp = Server.Createobject("MSXML2.ServerXMLHTTP")
    xmlHttp.Open "GET", strUrl, False
    xmlHttp.setRequestHeader "User-Agent", "asp httprequest"
    xmlHttp.setRequestHeader "content-type", "application/x-www-form-urlencoded"
    xmlHttp.Send
    response.write xmlHttp.responseText
    xmlHttp.abort()
    set xmlHttp = Nothing   
end function 

DIM content_type: content_type=Request.ServerVariables("HTTP_ACCEPT")
'IF (INSTR(UCASE(content_type),"JSON")>0) THEN
    Response.ContentType = "application/json"
    Response.CharSet = "ISO-8859-1"
    AsyncCall "https://server.panax.io:8081/startSQL"

%>
<% 'END IF %>