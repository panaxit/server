<!--#include file="vbscript.asp"-->
<%
    Response.ContentType = "application/json" 
    Response.Status = "401 Unauthorized"
    set rsResult = login
    IF Err.number<>0 Then
        ErrorDesc=""
        IF session("user_login")<>"" THEN
            ErrorDesc=SqlRegEx.Replace(Err.Description, "")
        ELSE 
            ErrorDesc="Conexión no autorizada"
        END IF
        Session.Contents.Remove("StrCnn")
        %>
	    {
	    "code": 2
	    , "success": false
	    , "status": "<%= session("status") %>"
        , "user_login": "<%= session("user_login") %>"
        , "database_id": "<%= session("database_id") %>"
        , "connection_id": "<%= session("connection_id") %>"
	    , "message": `<%= REPLACE(REPLACE(SqlRegEx.Replace(ErrorDesc, ""),"\","\\"),CHR(13),"\n") %>`
	    }
    <% 	response.end
    End If
    If rsResult.BOF and rsResult.EOF Then
	    Session("AccessGranted") = FALSE
        session("status") = "unauthorized"
    ELSE
        Response.Status = "200 Ok"
	    Session("AccessGranted") = TRUE
        session("status") = "authorized"
	    Response.Cookies("AntiPopUps") = REQUEST.FORM("AntiPopUps")
	    Response.Cookies("AntiPopUps").Expires = Date() + 1
	    Session.Timeout = 600
	    session("user_id")=rsResult(0)
        session("expires") = DateAdd("n", session.Timeout, NOW)
        IF session("user_id")="1" THEN
            session("debug") = TRUE
        END IF

    	Dim oCn: Set oCn = Server.CreateObject("ADODB.Connection")
	    oCn.ConnectionTimeout = 5
	    oCn.CommandTimeout = 60
        DIM StrCnn: StrCnn = "driver={SQL Server};server="&SESSION("secret_server_id")&";uid="&SESSION("secret_database_user")&";pwd="&SESSION("secret_database_password")&";database="&SESSION("secret_database_name")
        If oCn.State = 0 THEN
            ON ERROR RESUME NEXT
            oCn.Open StrCnn
        END IF

	    oCn.execute "IF EXISTS(SELECT 1 FROM INFORMATION_SCHEMA.ROUTINES IST WHERE routine_schema IN ('$Application') AND ROUTINE_NAME IN ('OnStartUp')) BEGIN EXEC [$Application].OnStartUp END"
    %>
    	    {
	    "success": true
        , "userId": "<%= session("user_id") %>"
        , "user_login": "<%= session("user_login") %>"
        , "database_id": "<%= session("database_id") %>"
        , "connection_id": "<%= session("connection_id") %>"
    <%
    FOR EACH oField IN rsResult.fields %>
    <% IF oField.name="" Then %>
    <%
    %>
    <% END IF %>
    <% IF TypeName(oField)="Field" THEN %><% sType=TypeName(oField.value): sValue=oField.value %><% ELSE %><% sType=TypeName(oField): sValue=oField %><% END IF %>
    <% SESSION(oField.name)= sValue %>
		        ,"<%= oField.name %>":<% SELECT CASE UCASE(sType): CASE "NULL": %>null<% CASE "BOOLEAN": %><% IF sValue THEN %>true<% ELSE %>false<% END IF %><% CASE ELSE %>"<%= RTRIM(REPLACE(replaceMatch(sValue, "["&chr(13)&""&chr(10)&""&vbcr&""&vbcrlf&"]", ""&vbcrlf),"""", """")) %>"<% END SELECT %> 
	        <% NEXT %>
    }
    <% rsResult.Close 
    END IF%>
    <% IF NOT(Session("AccessGranted")) THEN %>
	    {
	    "success": false
        , "user_login": "<%= session("user_login") %>"
        , "status": "unauthorized"
	    , "message": "Nombre de usuario o contraseña inválidos"
        , "source": "<%= REPLACE(strSQL,"""","""") %>"
	    }
    <% END IF 
    If oCn.State = 1 THEN
        oCn.Close
    END IF
%>