﻿<!--#include file="vbscript.asp"-->
<% 'Response.CodePage = 65001
'Response.CharSet = "UTF-8"
Set oCn = Server.CreateObject("ADODB.Connection")
oCn.ConnectionTimeout = 5
oCn.CommandTimeout = 60

Response.ContentType = "application/json"
Response.CharSet = "ISO-8859-1"
Dim key, value
Dim RegEx: Set RegEx = New RegExp
With RegEx
    .Pattern = "(\[[^\[]*\])+"
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
End With
IF INSTR(Request.ServerVariables("HTTP_ACCEPT"), "text/html") AND Session("user_login") = "uriel@panax.io" THEN
    Response.ContentType = "text/html"
    for each x in Request.ServerVariables
      response.write("<!--" & x & ": " & Request.ServerVariables(x) & "-->" & vbcrlf)
    next
END IF
IF Request.Form.Count>0 THEN 
    FOR EACH key IN Request.Form
        value=Request.Form(key)
        IF (key="user_login" OR key="user_id" OR key="database_id") AND value<>Session(key) THEN
            Session.Abandon
            Response.AddHeader "StatusMessage", "Unauthorized"
            Response.Status = "401 Unauthorized"
            Response.end
        ELSE
            IF value="null" THEN
                Session(key) = null
            ELSE
                session(key) = Request.Form(key)
            END IF
        END IF
    NEXT %>{"success":true}<%
END IF
checkConnection(oCn)
If oCn.State = 1 THEN
    oCn.Close
END IF
%>{ 
"userId": "<%= session("user_id") %>"
, "user_login": "<%= TRIM(Session("user_login")) %>"
, "referer": "<%= request.serverVariables("referer") %>"
, "connection_id": "<%= SESSION("connection_id") %>"
<%  Dim session_name
For Each session_name in Session.Contents 
    IF NOT(TypeName(Session.Contents(session_name))="DOMDocument" or TypeName(Session.Contents(session_name))="Null" or TypeName(Session.Contents(session_name))="Nothing" or session_name="StrCnn" or session_name="debug" or session_name="AccessGranted" or INSTR(session_name,"secret_")>0 or session_name="connection_id") THEN %>, "<%= session_name %>": "<%= TRIM(RegEx_JS_Escape.Replace(Session.Contents(session_name), "\$&")) %>" 
<%      END IF
Next %>
<% IF session("user_id")<>"" THEN  %>
    , "expires": "<%= DateAdd("n", session.Timeout, NOW) %>"
    <% END IF %>
}