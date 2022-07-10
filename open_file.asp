<!--#include file="vbscript.asp"-->
<%
DIM oConfiguration:	set oConfiguration = Server.CreateObject("MSXML2.DOMDocument"): 
oConfiguration.Async = false: 
oConfiguration.setProperty "SelectionLanguage", "XPath"
oConfiguration.Load(Server.MapPath("../../../config/system.config"))

DIM content_type: content_type=Request.ServerVariables("HTTP_ACCEPT")
IF NOT(Session("AccessGranted")) THEN 
    Response.CharSet = "ISO-8859-1"
    Response.Status = "401 Unauthorized"
    IF (INSTR(UCASE(content_type),"JSON")>0) THEN
    Response.ContentType = "application/json" %>
    {
        "status":"unauthorized"
      , "message":""
    }
    <% 
    ELSE %>
<html><body>Usuario no autorizado</body></html>
<%  END IF
    response.end
END IF
DIM StrCnn: StrCnn = "driver={SQL Server};server="&SESSION("secret_server_id")&";uid="&SESSION("secret_database_user")&";pwd="&SESSION("secret_database_password")&";database="&SESSION("secret_database_name")
Set cn = Server.CreateObject("ADODB.Connection")
cn.ConnectionTimeout = 0
cn.CommandTimeout = 0
cn.Open StrCnn
IF Err.Number<>0 THEN
    Response.ContentType = "application/javascript"
    Response.CharSet = "ISO-8859-1"
    Response.Clear()
    Response.Status = "401 Unauthorized"
	ErrorDesc=Err.Description
    %>
    this.status='exception';
    this.statusType='unauthorized';
    this.message="<%= REPLACE(REPLACE(ErrorDesc, "[Microsoft][ODBC SQL Server Driver][SQL Server]", ""), """", "\""") %>";
    <%
    response.end
END IF
cn.Execute("SET LANGUAGE SPANISH")

Dim RegEx: Set RegEx = New RegExp
With RegEx
    .Pattern = "(\[[^\[]*\])+"
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
End With

Dim objConn, strFile
Dim intCampaignRecipientID

strRFC = Request.QueryString("rfc")
strFullPath = REPLACE(Request.QueryString("full_path"),"/","\")
strFile = Right(strFullPath, InStr(StrReverse(strFullPath),"\")-1)

'base_folder="D:\Dropbox (Personal)\Proyectos\Petro\FilesRepository"
DIM xRepository: SET xRepository=oConfiguration.documentElement.selectSingleNode("/configuration/Repositories/Folder[@Id='upload']")
IF NOT (xRepository IS NOTHING) THEN
    base_folder=xRepository.getAttribute("Location")
ELSE
    base_folder=server.MapPath("../../../FilesRepository")
END IF
    
IF INSTR(UCASE(strFullPath),UCASE("C:\fakepath"))>0 THEN
    strFullPath=base_folder & "\" & strFile
ELSEIF INSTR(strFullPath,"C:\")<=0 AND INSTR(strFullPath,"D:\")<=0 THEN
    'TODO: Hacerlo genérico
    strFullPath=base_folder & "\SociosDeNegocios\Autorizado\"&strRFC & "\" & strFile
END IF

Dim objStream
DIM xslFile
DIM rsResult

If strFile <> "" Then
    Response.Buffer = False
    IF INSTR(UCASE(strFile),".PDF")>0 THEN
        Set objStream = Server.CreateObject("ADODB.Stream")
        objStream.Type = 1 
        objStream.Open
        objStream.LoadFromFile(strFullPath)
    
        Response.ContentType = "application/pdf"
        Response.Addheader "Content-Disposition", "inline; filename=" & strFile
        Response.BinaryWrite(objStream.Read)
        objStream.Close
        Set objStream = Nothing
    ELSEIF INSTR(UCASE(strFile),".XML")>0 THEN
        DIM xmlDoc:	set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
        xmlDoc.Async = false
        'xmlDoc.setProperty "SelectionNamespaces", "xmlns:cfdi='http://www.sat.gob.mx/cfd/3'"
        xmlDoc.SetProperty "SelectionLanguage", "XPath"
        xmlDoc.Load(strFullPath)
        IF xmlDoc.selectNodes("//*[namespace-uri()='http://www.sat.gob.mx/cfd/3']").length>0 THEN
        DIM command: command="BEGIN TRY DECLARE @Comprobante [xml]; SELECT @Comprobante='"&xmlDoc.xml&"' SET NOCOUNT ON; EXEC #SAT.validarFactura @Comprobante=@Comprobante OUTPUT; SELECT @Comprobante; END TRY BEGIN CATCH DECLARE @Message NVARCHAR(MAX); SELECT @Message=ERROR_MESSAGE(); EXEC [$Table].[getCustomMessage] @Message=@Message, @Exec=1; END CATCH"
            SET rsResult = cn.Execute(command)
            xmlDoc.LoadXML(rsResult(0))

            Response.CharSet = "ISO-8859-1"
            response.ContentType = "text/html" 
            xslFile=server.MapPath("..\resources\factura.xslt")
            Set xslDoc=Server.CreateObject("Microsoft.XMLDOM")
            xslDoc.async="false"
            xslDoc.load(xslFile)
            xmlDoc.loadXML(xmlDoc.transformNode(xslDoc))
        ELSEIF xmlDoc.selectNodes("//*[namespace-uri()='http://www.sat.gob.mx/cfd/3']").length>0 THEN
            Response.CharSet = "ISO-8859-1"
            response.ContentType = "text/html" 
            xslFile=server.MapPath("..\resources\aviso.xslt")
            Set xslDoc=Server.CreateObject("Microsoft.XMLDOM")
            xslDoc.async="false"
            xslDoc.load(xslFile)
            xmlDoc.loadXML(xmlDoc.transformNode(xslDoc))
        ELSE
            response.CodePage = 65001
            response.CharSet = "UTF-8"
            response.ContentType = "text/xml" 
        END IF

        If (xmlDoc.parseError.errorCode <> 0) Then  
            Set myErr = xmlDoc.parseError  
            response.write myErr.reason
            response.end
        End If  

        'response.write "<!DOCTYPE html>"
        response.write xmlDoc.xml

    ELSEIF INSTR(UCASE(strFile),".TXT")>0 THEN
        response.ContentType = "text/html" 
        Response.CharSet = "ISO-8859-1"
        dim fs,f
        set fs=Server.CreateObject("Scripting.FileSystemObject")
        set f=fs.OpenTextFile(strFullPath,1,false)
        f.Close
        set f=Nothing
        set fs=Nothing
    ELSE
        Set objStream = Server.CreateObject("ADODB.Stream")
        objStream.Type = 1 
        objStream.Open
        objStream.LoadFromFile(strFullPath)
        Response.ContentType = "application/x-unknown"
        Response.Addheader "Content-Disposition", "attachment; filename=" & strFile
        Response.BinaryWrite objStream.Read
        objStream.Close
        Set objStream = Nothing
    END IF
End If
%>
