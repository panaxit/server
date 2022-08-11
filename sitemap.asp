<% SERVER.SCRIPTTIMEOUT = 4800 %>
<%
DIM content_type: content_type=Request.ServerVariables("HTTP_ACCEPT")
Server.ScriptTimeOut=1200
response.Buffer=true
DIM StrCnn: StrCnn = "driver={SQL Server};server="&SESSION("secret_server_id")&";database="&SESSION("secret_database_name")&";uid="&SESSION("secret_database_user")&";pwd="&SESSION("secret_database_password")
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

DIM oCn:	set oCn=server.createobject("adodb.connection")
oCn.ConnectionTimeout = 5
oCn.CommandTimeout = 60
ON ERROR RESUME NEXT
oCn.Open StrCnn
IF Err.Number<>0 THEN 
    Response.ContentType = "application/json"
    Response.CharSet = "ISO-8859-1"
    Response.Status = "401 Unauthorized" %>
	{
	"success": false,
	"message": "No se pudo establecer una conexión con la base de datos <%= sDatabaseName %>: <%= REPLACE(RegEx.Replace(Err.Description, ""),"""","""") %>"
	}
<% 	response.end
END IF
oCn.Execute("SET LANGUAGE SPANISH")

DIM rsRecordSet:	Set rsRecordSet = Server.CreateObject("ADODB.RecordSet")
rsRecordSet.CursorLocation 	= 3
rsRecordSet.CursorType 		= 3

DIM oSiteMap:	
set oSiteMap = Server.CreateObject("Microsoft.XMLDOM")
oSiteMap.Async = false:
'IF NOT(SESSION("UserSiteMap") IS NOTHING) THEN
'    set oSiteMap =SESSION("UserSiteMap")
'ELSE
        Response.CharSet = "ISO-8859-1"
		DIM sSQL:	sSQL="EXEC [$Security].UserSitemap @@user_id=-1"'&session("user_id")
	    DIM sitemapFile, sitemap
        sitemapFile=server.MapPath(".")&"\..\web.sitemap"
        set fso=CreateObject("Scripting.FileSystemObject")
        IF fso.FileExists(sitemapFile) THEN
            Set sitemap=Server.CreateObject("Microsoft.XMLDOM")
            sitemap.async="false"
            sitemap.load(sitemapFile)
            sSQL = sSQL & ", @sitemap='"& replace(sitemap.xml,"'","''") &"'"
        End If
		IF session("lang")<>"" THEN sSQL = sSQL&", @lang="&session("lang") END IF
		set rsRecordSet=oCn.execute(sSQL)
		SELECT CASE Err.Number
		CASE -2147217900 %>
			{
			success: false,
			message: "Error: <%= REPLACE(Err.Description, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "") %><% IF session("user_id")=-1 THEN response.write Err.Description & " \n\n"&sSQL %>"
			}
<%			CASE 0
			'continue
		CASE ELSE %>
			{
			success: false,
			message: "Error <%= Err.Number %>!, no se pudo recuperar la información.<% IF session("user_id")=1 THEN response.write Err.Description&" <br><br>"&sSQL %>"
			}
<%				response.end
			Err.Clear
		END SELECT
	ON ERROR  GOTO 0
    'response.write sSQL: response.end
	'response.write rsRecordSet(0): response.end
    IF SESSION("user_login")="webmaster" or SESSION("debug") THEN
        response.write "<!--"&REPLACE(sSQL,"--","- -")&"-->"
        response.write "<!--"&REPLACE(sitemapFile,"--","- -")&"-->"
    END IF %>
	<% oSiteMap.LoadXML(rsRecordSet(0)) %>
	<% SET SESSION("UserSiteMap")=oSiteMap 
'END IF %>
<% 
    Response.CodePage = 65001
    Response.CharSet = "UTF-8"
    Response.ContentType = "text/xml"
    'Response.write "<!-- "&sSQL&"-->"
    response.Write "<?xml-stylesheet type=""text/xsl"" href=""sitemap.xslt"" role=""sitemap"" target=""@#shell @#sitemap"" action=""replace""?>"
    response.write oSiteMap.xml %>
<%'= transformXML(SESSION("UserSiteMap"), server.MapPath("..\templates\mapSite.xsl")) %>
<% If oCn.State = 1 THEN
    oCn.Close
END IF %>
<% SET rsRecordSet = NOTHING %>
<% SET oCn = NOTHING %>
