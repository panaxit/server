<%
DIM StrCnn: StrCnn = "driver={SQL Server};server="&SESSION("secret_server_id")&";uid="&SESSION("secret_database_user")&";pwd="&SESSION("secret_database_password")&";database="&SESSION("secret_database_name")
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
ON ERROR RESUME NEXT
Set cn = Server.CreateObject("ADODB.Connection")
cn.ConnectionTimeout = 0
cn.CommandTimeout = 0
cn.Open StrCnn
IF Err.Number<>0 THEN
    Response.ContentType = "text/html"
    Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
    Response.Clear()
    Response.Status = "401 Unauthorized"
	ErrorDesc=Err.Description
    %>
    this.status='exception';
    this.statusType='unauthorized';
    this.message="<%= REPLACE(REPLACE(ErrorDesc, "[Microsoft][ODBC SQL Server Driver][SQL Server]", ""), """", """") %>";
    <%
    response.end
END IF
cn.Execute("SET LANGUAGE SPANISH")
%>
<!--#include file="vbscript.asp"-->
<%
'Response.CodePage = 65001
'Response.CharSet = "UTF-8"
ON ERROR RESUME NEXT
strSQL = request.querystring("RoutineName") & " " & request.querystring("Parameters")

session_id=request.querystring("sessionId")
IF session_id="" THEN
    session_id="user_"&session("user_id") &""
END IF

DIM file_location
file_location=server.MapPath(".")&"\..\..\sessions\"&session_id&".xml"
DIM xmlDoc:	set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = false
xmlDoc.Load(file_location)

response.CodePage = 65001
response.CharSet = "UTF-8"
response.ContentType = "text/xml" 
if (xmlDoc.selectSingleNode("//x:source/*") is nothing) then
    response.write xmlDoc.xml
else
    response.write xmlDoc.selectSingleNode("//x:source/*").xml
end if
%>