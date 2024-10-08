﻿<!--#include file="vbscript.asp"-->
<% 
'for each x in Request.ServerVariables
  'response.write("<!--" & x & ": " & Request.ServerVariables(x) & "-->")
'next
'response.write("<!--" & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING") & "-->")
DIM content_type: content_type=Request.ServerVariables("HTTP_ACCEPT")
IF content_type="*/*" OR content_type="*/*, */*" THEN
    content_type="text/xml"
END IF
'IF content_type="" THEN
'    response.write content_type
'    response.end
'END IF
DIM authorization: authorization = Request.ServerVariables("HTTP_AUTHORIZATION")
If (authorization<>"") Then
    login
End if

Server.ScriptTimeOut=1200
response.Buffer=true
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

Dim RegEx: Set RegEx = New RegExp
With RegEx
    .Pattern = "'?(\.?\[[^\[]*\])+'?"
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
End With

Function getQuerystring(keys, join, bFormatValue)
    Dim string: string=""
    Dim key
    For Each key in keys
        If Request.QueryString(key).count > 0 Then
            Dim item, value
            If Request.QueryString(key).count > 1 Then
                For Each item In Request.QueryString(key)
                    value = URLDecode(item)
                    If (bFormatValue) Then
                        value = FormatValue(value)
                    End If
                    string = string & join & value
                Next
            Else
                value = URLDecode(Request.QueryString(key))
                If (bFormatValue) Then
                    value = FormatValue(value)
                End If
                string = string & join & value 'Ya es decodificado automáticamente. Caso "folio like '%2304'", lo decodifica nuevamente como "folio like '#04'"
            End If
            'IF join<>"" THEN
            '    string = replace(string, join, "", 1, 1)
            'END IF
        End If
    Next
    getQuerystring = string
End Function

Sub manageError(Err)
    If TypeName(Err) = "String" Then
        Description = Err
    ELSE 
        Description = Err.Description
    End If

    Response.CharSet = "UTF-8"
    Response.Clear()
    IF INSTR(Description, "interbloqueo") THEN
        Response.Status = "423 Locked"
    ELSE 
        Response.Status = "409 Conflict"
    END IF

    DIM message: message=RegEx.Replace(Description, "")
    IF message="" AND r>max_recordsets THEN
        message = "La solicitud devolvió más conjuntos de datos de los permitidos"
    ELSEIF INSTR(message,"SQL Server does not exist or access denied")>0 OR INSTR(message,"Communication link failure")>0 THEN
        Response.Status = "503 Service Unavailable" '"408 Request Timeout"
        AsyncCall "http://localhost:8080/startSQL"
    'ELSEIF INSTR(UCASE(message), UCASE("clave duplicada"))>0 THEN
	'	message="PRECAUCIÓN: No se puede insertar un registro duplicado."
	'ELSEIF INSTR(UCASE(message), UCASE("La columna no admite valores NULL"))>0 THEN
	'	message="El campo no se puede quedar vacío"
	'ELSE
		'message="El sistema no pudo completar el proceso y envió el siguiente mensaje: \n\n"&message
	END IF

    IF INSTR(content_type,"xml")>0 THEN 
        response.ContentType = "text/xml" 
        IF SESSION("user_login")="webmaster" OR INSTR(SESSION("user_login"),"@panax.io")<>0 OR SESSION("debug") THEN
            response.write "<!--"&strSQL&"-->"
        END IF
    %>
<?xml-stylesheet type="text/xsl" href="message.xslt" role="message" target="body" action="append"?>
<xo:message xmlns:xo="http://panax.io/xover" xo:id="message_<%= REPLACE(REPLACE(REPLACE(NOW(),":",""),"/","")," ","_") %>" type="exception"><%= REPLACE(REPLACE(message,">","&gt;"),"<","&lt;") %></xo:message>
<%  ELSEIF INSTR(content_type,"json")>0 THEN
    Response.ContentType = "application/json" %>
//<%= strSQL  %>
<%  ELSE 
    Response.ContentType = "application/json" %>
    {"message":"<%= REPLACE(message, """", "\""") %>"
    <%  IF 1=1 OR session("debug")=TRUE THEN %>
    , "source": "<%= REPLACE(strSQL, """", "\""") %>"}
    <%  END IF 
    END IF
    ON ERROR GOTO 0
End Sub

ON ERROR RESUME NEXT
DIM StrCnn: StrCnn = "driver={SQL Server};server="&SESSION("secret_server_id")&";uid="&SESSION("secret_database_user")&";pwd="&SESSION("secret_database_password")&";database="&SESSION("secret_database_name")
Set oCn = Server.CreateObject("ADODB.Connection")
oCn.ConnectionTimeout = 5
oCn.CommandTimeout = 0
oCn.Open StrCnn
IF Err.Number<>0 THEN
    Response.ContentType = "application/json"
    Response.CharSet = "UTF-8"
    Response.Clear()

	ErrorDesc=RegEx.Replace(Err.Description, "")
    IF INSTR(ErrorDesc,"SQL Server does not exist or access denied")>0 OR INSTR(ErrorDesc,"Communication link failure")>0 THEN
        Response.Status = "503 Service Unavailable" '"408 Request Timeout"
        AsyncCall "http://localhost:8080/startSQL"
        'asyncCall "reconnect.asp"
    ELSE
        Response.Status = "401 Unauthorized"
    END IF
    IF ErrorDesc<>"" THEN
    %>{"message":"<%= REPLACE(ErrorDesc, """", "\""") %>"
    <%  IF 1=1 OR session("debug")=TRUE THEN %>
    , "source": "<%= REPLACE(strSQL, """", "\""") %>"}
    <%  END IF 
    END IF
    response.end
END IF
oCn.Execute("SET LANGUAGE SPANISH")
ON ERROR GOTO 0
    %>
<%
DIM rebuild: rebuild=eval(Request.ServerVariables("HTTP_X_REBUILD"))
DIM debug: debug=eval(Request.ServerVariables("HTTP_X_DEBUGGING"))
IF debug="" THEN
    debug = session("debug")
END IF
IF debug="" THEN
    debug=FALSE
END IF
SESSION("debug") = debug
    
Response.CharSet = "ISO-8859-1"
DIM api_key: api_key = Request.ServerVariables("HTTP_API_KEY") 'TODO: Implement
DIM root_node: root_node = Request.ServerVariables("HTTP_X_ROOT_NODE")
IF root_node="" THEN
    root_node="xo:response"
END IF
DIM page_size: page_size = Request.ServerVariables("HTTP_X_PAGE_SIZE")
IF page_size="" THEN
    page_size="100000"
END IF
DIM page_index: page_index = Request.ServerVariables("HTTP_X_PAGE_INDEX")
IF page_index="" THEN
    page_index="1"
END IF
DIM max_records: max_records = Request.ServerVariables("HTTP_X_MAX_RECORDS")
DIM namespaces: namespaces = Request.ServerVariables("HTTP_X_NAMESPACES")
'IF 1=1 OR Request.ServerVariables("HTTP_ROOT_NODE").Count>0 AND root_node = "" THEN

DIM output_parameter: output_parameter = Request.ServerVariables("HTTP_X_OUTPUT_PARAMETER")
IF output_parameter="" THEN
    output_parameter=""
END IF

DIM max_recordsets: max_recordsets = Request.ServerVariables("HTTP_X_MAX_RECORDSETS")

IF max_recordsets="" THEN
    max_recordsets = 1
END IF

'RESOURCE
DIM sType
sType = Request.ServerVariables("HTTP_QUERY_TYPE")
DIM command: 
command = getQuerystring(Array("command","FROM"), ",", False)
command = MID(command, 2)
IF command="" THEN
    manageError("No command provided")
END IF

ON ERROR RESUME NEXT
'DIM sRequestType: sRequestType="SET NOCOUNT ON; IF OBJECT_ID('#Object.FindObjectsInQuery') IS NOT NULL BEGIN SELECT TOP 1 [Type], [Object_Name] FROM #Object.FindObjectsInQuery('"&REPLACE(command,"'","''")&"') ORDER by Position END ELSE BEGIN SELECT [Type], [Object_Name]=QUOTENAME(OBJECT_SCHEMA_NAME(o.object_id))+'.'+QUOTENAME(OBJECT_NAME(o.object_id)) FROM sys.objects o WHERE o.object_id=OBJECT_ID('"&command&"') END"
DIM sRequestType: sRequestType="SET NOCOUNT ON; IF OBJECT_ID('#panax.getObjectInfoForUser') IS NOT NULL BEGIN SELECT TOP 1 [Type], [Object_Name] FROM #panax.getObjectInfoForUser('"&REPLACE(command,"'","''")&"','"&SESSION("user_login")&"') o END ELSE BEGIN IF OBJECT_ID('#Object.FindObjectsInQuery') IS NOT NULL BEGIN SELECT TOP 1 [Type], [Object_Name] FROM #Object.FindObjectsInQuery('"&REPLACE(command,"'","''")&"') ORDER by Position END ELSE BEGIN SELECT [Type], [Object_Name]=QUOTENAME(OBJECT_SCHEMA_NAME(o.object_id))+'.'+QUOTENAME(OBJECT_NAME(o.object_id)) FROM sys.objects o WHERE o.object_id=OBJECT_ID('"&REPLACE(command,"'","''")&"') END END"
'response.write "<!-- "&sRequestType&" -->": response.end
'strSQL=URLDecode(sRequestType) 'El símbol de (+) %2B es decodificado mal, revisar si es necesario decodificar
DIM rsType: SET rsType = oCn.Execute(sRequestType)
DIM sRoutineName: sRoutineName = URLDecode(request.querystring("RoutineName"))
IF NOT (rsType.BOF and rsType.EOF) THEN 
    sType = rsType("Type")
    sRoutineName = rsType("Object_Name")
ELSEIF Request.ServerVariables("HTTP_QUERY_TYPE")<>"" THEN
    sType = Request.ServerVariables("HTTP_QUERY_TYPE")
ELSE
    Response.Status = "404 Not found"
    Response.End
END IF

'response.write "sRoutineName: "&sRoutineName: response.end
IF INSTR(sType,"V")<>0 THEN
    sType = "T"
END IF    

'FIELDS
DIM data_fields: data_fields = Request.ServerVariables("HTTP_X_DATA_FIELDS")
data_fields = URLDecode(""&data_fields)

DIM data_value: data_value = Request.ServerVariables("HTTP_X_DATA_VALUE")
IF data_value<>"" THEN
    IF data_fields<>"" THEN data_fields=data_fields & ", " END IF
    data_fields = data_fields & "[value]=" & URLDecode(data_value)
END IF

DIM data_text: data_text = Request.ServerVariables("HTTP_X_DATA_TEXT")
IF data_text<>"" THEN
    IF data_fields<>"" THEN data_fields=data_fields & ", " END IF
    data_fields = data_fields & "[text]=" & URLDecode(data_text)
ELSEIF data_value<>"" THEN
    data_fields = data_fields & "[text]=" & URLDecode(data_value)
END IF
IF request.querystring("fields")<>"" THEN
    IF data_fields<>"" THEN data_fields = data_fields & ", " END IF
    data_fields = data_fields & URLDecode(request.querystring("fields"))
END IF
IF data_fields="" THEN
    data_fields="*"
END IF

'PREDICATES (FOR TABLES AN FUNCTION TABLES)
DIM order_by: order_by = Request.ServerVariables("HTTP_X_ORDER_BY")
IF (order_by="") THEN
    order_by="(SELECT NULL)"
END IF
DIM data_predicate: data_predicate = Request.ServerVariables("HTTP_X_DATA_PREDICATE")
IF INSTR(sType,"T")<>0 THEN
    IF data_predicate="" THEN
        data_predicate = Request.ServerVariables("HTTP_X_DATA_FILTERS")
    END IF
    IF data_predicate="" THEN
        data_predicate = URLDecode(request.querystring("predicate"))
    END IF
    IF data_predicate="" THEN
        data_predicate = URLDecode(request.querystring("filters"))
    END IF
    data_predicate = data_predicate & getQuerystring(Array("WHERE"), " AND ", False)
    extra_predicate= extra_predicate & getQuerystring(Array("AND"), " AND ", False)
    extra_predicate= extra_predicate & getQuerystring(Array("OR"), " OR ", False)
END IF
IF max_records<>"" THEN
    max_records_predicate = "WHERE [@meta:resultCount] <= " & max_records
END IF

'data_predicate = URLDecode(data_predicate) 'Se comenta porque ya se decodifica en las funciones y se podría decodificar dos veces indebidamente
DIM payload
set xmlParameters = Server.CreateObject("Microsoft.XMLDOM"): 
xmlParameters.Async = false: 
xmlParameters.setProperty "SelectionLanguage", "XPath"
call xmlParameters.setProperty("SelectionNamespaces", "xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'")

' PARAMETERS
DIM sParameters
'sParameters=replaceMatch(URLDecode(command),"^"&replaceMatch(sRoutineName,"([\[\]\(\)\.\$\^])","\$1")&"\s*\(?|\)$","")
If Request.TotalBytes > 0 Then
    DIM payload_parameter_name: payload_parameter_name=Request.ServerVariables("HTTP_X_PAYLOAD_PARAMETER_NAME")
    DIM dataType: dataType="string"
    IF INSTR(Request.ServerVariables("HTTP_CONTENT_TYPE"),"xml")>0 THEN
        DIM xPayload
        Set xPayload=Server.CreateObject("Microsoft.XMLDOM")
        xPayload.async="false"
        xPayload.load(request)
        xPayload.selectNodes("//comment()").removeAll()
        payload = URLDecode(xPayload.xml)
        dataType="xml"
    ELSE
        payload=BytesToStr(Request.BinaryRead(Request.TotalBytes))
    END IF
    IF INSTR(sType,"P")<>0 OR INSTR(sType,"F")<>0 THEN
        IF (payload_parameter_name<>"") THEN
            payload_parameter_name=" name="""&payload_parameter_name&""""
        END IF
        xmlParameters.LoadXML("<parameters><param"&payload_parameter_name&" dataType="""&dataType&"""><![CDATA["&payload&"]]></param></parameters>")
    ELSEIF INSTR(sType,"T")<>0 THEN
        IF (payload_parameter_name<>"") THEN
           data_predicate = data_predicate & payload_parameter_name&"='"&REPLACE(payload,"'","''")&"'"
        ELSE
           data_predicate = data_predicate & payload
        END IF
    END IF
End If
IF xmlParameters.documentElement IS NOTHING THEN
    xmlParameters.LoadXML("<parameters/>")
END IF

'CACHÉ
DIM full_request: full_request=Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")'data_fields&"&"&command&"&"&data_predicate
DIM file_location, alt_file_location, file_name
IF INSTR(content_type,"xml")>0 THEN
    file_name=Hash("md5",REPLACE(full_request,"ñ","")) &".xml"
        'response.Clear
ELSEIF INSTR(content_type,"javascript")>0 THEN
    file_name=Hash("md5",REPLACE(full_request,"ñ","")) &".js"
END IF

'response.write "full_request: "&file_name: response.end
ON ERROR RESUME NEXT
DIM parent_folder: parent_folder=server.MapPath("\")&"\cache\"&SESSION("secret_database_user")&"\"
DIM alt_parent_folder: alt_parent_folder=server.MapPath("\")&"\cache\"&SESSION("user_login")&"\"
set fso=CreateObject("Scripting.FileSystemObject")
If  Not fso.FolderExists(parent_folder) Then      
    CreateFolder parent_folder
End If
file_location=parent_folder&file_name

If  Not fso.FolderExists(alt_parent_folder) Then      
    CreateFolder alt_parent_folder
End If
alt_file_location=alt_parent_folder&file_name

IF Err.Number<>0 THEN
    manageError(Err)
    response.end
END IF
ON ERROR GOTO 0

DIM oXMLFile:	set oXMLFile = Server.CreateObject("Microsoft.XMLDOM")
oXMLFile.Async = false

IF fso.FileExists(file_location) THEN
    Set file = fso.GetFile(file_location)
    Dim fileAge
    fileAge = Now() - file.DateLastModified
    Dim maxAge
    maxAge = -1 ' Default value for max-age
    If Request.ServerVariables("HTTP_CACHE_CONTROL") <> "" Then
        Dim cacheControl
        cacheControl = Request.ServerVariables("HTTP_CACHE_CONTROL")
        Dim maxAgeIndex
        maxAgeIndex = InStr(cacheControl, "max-age=")
        If maxAgeIndex > 0 Then
            maxAge = CInt(substringBefore(Mid(cacheControl, maxAgeIndex + 8),","))
        End If
    End If
    If maxAge >= 0 And fileAge < (maxAge / 86400) Then ' Convert max-age from seconds to days
        ' Return the cached version of the file
        Response.CacheControl = "public, max-age=" & maxAge
        oXMLFile.load(file_location)
        Response.CodePage = 65001
        Response.CharSet = "UTF-8"
        Response.ContentType = "text/xml"
        Response.write "<!-- Desde cache: "&file_name&". "&full_request&"-->"
        DIM xslFile, xslValues
        xslFile=server.MapPath(".")&"\normalize_values.xslt"
        Set xslValues=Server.CreateObject("Microsoft.XMLDOM")
        xslValues.async="false"
        xslValues.load(xslFile)
        oXMLFile.loadXML(oXMLFile.transformNode(xslValues))
        Response.Write oXMLFile.xml
        Response.end
    End If
END IF

IF (INSTR(sType,"P")<>0 OR INSTR(sType,"F")>0) THEN
    DIM sParamValue, bParameterString', aParameters
    'Set aParameters=Server.CreateObject("Scripting.Dictionary")

    DIM detect_input_variables: detect_input_variables = EVAL(Request.ServerVariables("HTTP_X_DETECT_INPUT_VARIABLES"))
    IF detect_input_variables="" THEN
        detect_input_variables = TRUE
    END IF
    DIM detect_output_variables: detect_output_variables = EVAL(Request.ServerVariables("HTTP_X_DETECT_OUTPUT_VARIABLES"))
    IF detect_output_variables="" THEN
        detect_output_variables = TRUE
    END IF
    DIM detect_missing_variables: detect_missing_variables = EVAL(Request.ServerVariables("HTTP_X_DETECT_MISSING_VARIABLES"))
    IF detect_missing_variables="" THEN
        detect_missing_variables = TRUE
    END IF

    IF request.querystring("Parameters")<>"" THEN
        sQueryParameters=URLDecode(request.querystring("Parameters"))
    END IF
    IF (detect_input_variables OR detect_output_variables) THEN
        'sQueryParameters=replaceMatch(URLDecode(command),"^"&replaceMatch(sRoutineName,"([\[\]\(\)\.\$\^])","\$1")&"\s*\(?|\)$","")
        command = sRoutineName
        IF detect_input_variables AND sQueryParameters<>"" THEN
            sQueryParameters=replaceMatch(sQueryParameters,"\bDEFAULT\b","'$&'")
            'response.write sQueryParameters: response.end

            sQueryParameters=replaceMatch(sQueryParameters, "^\(|\)$", "")
            'sQueryParameters=replaceMatch(sQueryParameters, "\@[^=]+=('|\d+|\w+|DEFAULT)", "'$&'")

            DIM sSQLXMLParams: sSQLXMLParams="SET NOCOUNT ON; IF OBJECT_ID('#panax.parameterStringToXML') IS NOT NULL BEGIN EXEC #panax.parameterStringToXML '"&REPLACE(sQueryParameters,"'","''")&"' END ELSE BEGIN SELECT CONVERT(XML,'') END"  
            IF debug THEN
                response.ContentType = "text/xml" 
                response.write "<!--"&sSQLXMLParams&"-->"
            END IF
            'response.write sSQLXMLParams: response.end
            ON ERROR RESUME NEXT
            SET rsParameters = oCn.Execute(sSQLXMLParams)
            IF Err.Number<>0 THEN
                manageError(Err)
                response.end
            END IF

            IF rsParameters.fields.Count>0 AND NOT(rsParameters.BOF AND rsParameters.EOF) THEN
                xmlParameters.LoadXML(rsParameters(0))
            END IF
        END IF
        
        'DIM sParameter, ns

        i=0

        DIM rsParameters
        DIM missingParameters
        'Set aOutputParameters=Server.CreateObject("Scripting.Dictionary")

        DIM rebuild_parameters_snippet
        IF rebuild THEN
            rebuild_parameters_snippet=", @rebuild=1"
        END IF
        DIM sSQLParams: sSQLParams="SET NOCOUNT ON; DECLARE @parameters XML; IF OBJECT_ID('[#panax].[getParameters]') IS NOT NULL BEGIN EXEC [#panax].[getParameters] '"&REPLACE(command,"'","''")&"', @parameters=@parameters OUT"&rebuild_parameters_snippet&"; END SELECT ISNULL(@parameters , '')"
        IF debug THEN
            response.ContentType = "text/xml" 
            response.write "<!--"&sSQLParams&"-->" & vbcrlf
            'response.end
        END IF
        ON ERROR RESUME NEXT
        SET rsParameters = oCn.Execute(sSQLParams)
        IF Err.Number<>0 THEN
            manageError(Err)
            response.end
        END IF
        DIM xmlOutputParameters:	set xmlOutputParameters = Server.CreateObject("Microsoft.XMLDOM"): xmlOutputParameters.Async = false: 
        DIM i, sOutputParams
        IF NOT(rsParameters.BOF AND rsParameters.EOF) AND rsParameters.fields.Count>0 THEN
	        xmlOutputParameters.LoadXML(rsParameters(0))
	        i=0
            IF xmlOutputParameters.documentElement IS NOTHING AND NOT(xmlParameters.selectSingleNode("parameters/*") IS NOTHING) THEN
                xmlOutputParameters.LoadXML(xmlParameters.xml)
	        END IF
        END IF
        FOR EACH sParameter IN request.querystring
	        IF testMatch(sParameter, "^\@") THEN
		        sParamValue=URLDecode(request.querystring(sParameter))
		        bParameterString=NOT(sParamValue="" OR UCASE(sParamValue)="NULL" OR UCASE(sParamValue)="DEFAULT" OR ISNUMERIC(sParamValue) OR testMatch(sParamValue, "^['@]"))
		        IF bParameterString THEN sParamValue="'"&REPLACE(sParamValue,"'","''")&"'" END IF
		        IF RTRIM(sParamValue)="" THEN sParamValue="NULL" END IF
                set param = xmlParameters.createElement("param")
                IF sParameter<>"@" THEN 'Allows unnamed parameters to be sent like @=param1_value, @=param2_value
                    param.setAttribute "name", sParameter
                END IF
                'param.setAttribute "value", sParamValue
                param.Text = sParamValue
                xmlParameters.selectSingleNode("parameters").appendChild(param)
	        END IF
        NEXT

	    IF NOT(xmlOutputParameters.documentElement IS NOTHING) THEN
		    DIM sParamsDeclaration
		    DIM sParamsDefinition
            DIM xParameter, sParameterName, oOtherNodes
            DIM sParameterType
		    FOR EACH oNode IN xmlOutputParameters.documentElement.selectNodes("/*/*")
                IF i>0 THEN
                    sParameters=sParameters&", "
                END IF
   			    i=i+1
                sParameterName = oNode.getAttribute("name")
                sParameterValue = oNode.text
                sParameterType = "string"
                set xParameter=xmlParameters.documentElement.selectSingleNode("/*/*[not(@name)][position()='"&i&"']|/*/*[translate(@name,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='"&UCASE(sParameterName)&"']")
                SET oOtherNodes = xmlParameters.documentElement.selectNodes("/*/*[@position>"&i&"]")
                'IF request.querystring(sParameterName).count > 0 THEN
                '    sParameterValue = request.querystring(sParameterName)
                'ELSE
                IF NOT(IsEmpty(xParameter) OR xParameter IS NOTHING) THEN
                    sParameterValue = xParameter.Text
                    IF NOT(IsEmpty(xParameter.getAttribute("xsi:type"))) THEN
                        sParameterType = xParameter.getAttribute("xsi:type")
                    END IF
                ELSEIF INSTR(sParameterName,"@@")=1 AND NOT(IsEmpty(SESSION(REPLACE("^"&sParameterName,"^@@","")))) THEN 'Los parámetros con doble arroba pueden mapear automáticamente a variables de sesión.
                    sParameterValue = SESSION(REPLACE(sParameterName,"@@",""))
                ELSE
                    sParameterValue = "DEFAULT"
                END IF                
                    
                IF oNode.getAttribute("isRequired")=1 AND oNode.getAttribute("isOutput")=0 AND ((IsEmpty(xParameter) OR sParameterType<>"string") AND sParameterValue="" OR sParameterValue="DEFAULT") THEN
                    missingParameters = TRUE
                    sParameterValue=""
                    oNode.setAttribute "missing", "true"
                ELSEIF sParameterValue="" AND (IsEmpty(xParameter) OR xParameter IS NOTHING) AND NOT(IsEmpty(oOtherNodes)) THEN
                    sParameterValue = "NULL"
                END IF

                oNode.setAttribute "value", ""
                oNode.text = sParameterValue
                IF NOT(INSTR(oNode.getAttribute("dataType"),"int")>0 OR INSTR(oNode.getAttribute("dataType"),"bit")>0) AND NOT(UCASE(sParameterValue)="NULL" OR sParameterValue="DEFAULT" OR getMatch(sParameterValue, "^'([\S\s]*)'$|^\(([\S\s]*)\)$").count>=1) THEN
                    sParameterValue = "'"&REPLACE(sParameterValue,"'","''")&"'"
                END IF
                IF INSTR(oNode.getAttribute("dataType"),"date")<>0 THEN
                    sParameterValue = replaceMatch(sParameterValue,"^(\d+)-(\d+)-(\d+)$","$1$2$3")
                END IF

                data_type=oNode.getAttribute("dataType")
                IF data_type="[decimal]" THEN
                    data_type="[decimal](10,5)"
                    sParameterValue = "NULL"
                END IF
                IF ISNULL(data_type) AND sParameterType="string" THEN
                    data_type = "nvarchar(MAX)"
                END IF
    			sParamsDeclaration=sParamsDeclaration& "DECLARE "&oNode.getAttribute("name")&" "&data_type&"; "
                IF oNode.getAttribute("isOutput")=0 AND sParameterValue="DEFAULT" THEN
                    sParameters=sParameters&"DEFAULT"
                ELSE
                    IF sParameterValue="DEFAULT" THEN
                        sParameterValue="NULL" 'Revisar si se debe iniciarlizar con el valor del default
                    END IF
                    sParamsDefinition=sParamsDefinition& "SELECT "&oNode.getAttribute("name")&"="&sParameterValue&";" 
                    'IF INSTR(sType,"P")<>0 THEN
                    '    sParameters=sParameters&sParameterName&"="&sParameterName
                    'ELSE
                        sParameters=sParameters&sParameterName
                    'END IF
                END IF
                IF oNode.getAttribute("isOutput")=1 THEN
                    sParameters=sParameters&" OUT"
			        IF sOutputParams<>"" THEN sOutputParams=sOutputParams&", " END IF
                    sOutputParams=sOutputParams& "["&REPLACE(oNode.getAttribute("name"), "@", "")&"]=" & oNode.getAttribute("name")
                END IF
		    NEXT
        END IF
    END IF
    'response.write xmlParameters.xml
    'response.write xmlOutputParameters.xml
    'response.end

    IF INSTR(sType,"F")<>0 THEN
        command = command & "(" & TRIM(sParameters) &")"
    ELSE
        command = command & " " & TRIM(sParameters)
    END IF
    'response.write xmlOutputParameters.xml: response.end
    IF missingParameters=TRUE AND detect_missing_variables=TRUE THEN 
        response.ContentType = "text/xml"
        Response.Status = "412 Precondition Failed" 
%>
<?xml-stylesheet type="text/xsl" href="prompt.xslt" role="modal" target="@#shell main" ?>
<xo:prompt xmlns:xo="http://panax.io/xover"><%= xmlOutputParameters.xml %></xo:prompt>
<%
        response.end
    END IF
ELSE
    command = sRoutineName
END IF 

FOR EACH sParameter IN request.querystring
	IF INSTR(sType,"T")<>0 AND NOT(testMatch(sParameter, "^\@|^FROM$|^AND$|^OR$|^WHERE$")) THEN
        'IF sParameters<>"" THEN
        '    sParameters=sParameters&", "
        'END IF
		'sParamValue=request.querystring(sParameter)
		'bParameterString=NOT(sParamValue="" OR sParamValue="NULL" OR sParamValue="DEFAULT" OR ISNUMERIC(sParamValue) OR 'testMatch(sParamValue, "^['@]"))
		'IF bParameterString THEN sParamValue="'"&REPLACE(sParamValue,"'","''")&"'" END IF
		'IF RTRIM(sParamValue)="" THEN sParamValue="NULL" END IF
        'sParameters=sParameters & sParameter&" IN ("&sParamValue&")"
        data_predicate = data_predicate & " AND "
        data_predicate = data_predicate & sParameter&" IN (NULL"&getQuerystring(Array(sParameter), ",", True)&")"
	END IF
NEXT

IF INSTR(sType,"T")<>0 THEN
    'data_fields="TOP 1000 "&data_fields&" "
    IF data_predicate<>"" THEN
        data_predicate=" WHERE 1=1 "&data_predicate
    END IF
    IF extra_predicate<>"" THEN
        extra_predicate=" WHERE 1=1 "&extra_predicate
    END IF
END IF 

IF INSTR(sType,"P")<>0 THEN
    strSQL="EXEC "&command &"; "
    IF sOutputParams<>"" THEN 
        strSQL=strSQL&"WITH XMLNAMESPACES('http://panax.io/xover' as xo, 'http://panax.io/state' as state, 'http://panax.io/metadata' as meta, 'http://panax.io/custom' as custom, 'http://panax.io/fetch/request' as source, 'http://www.mozilla.org/TransforMiix' as transformiix) SELECT (SELECT "&sOutputParams&" FOR XML PATH(''), TYPE) FOR XML PATH(''), ROOT('xo:parameters'), TYPE"
    END IF
ELSEIF INSTR(sType,"T")<>0 THEN 'Table  y Table Function
    IF namespaces<>"" THEN
        namespaces = ", " & namespaces
    END IF
    strSQL="SET NOCOUNT ON; SET TEXTSIZE 2147483647; DECLARE @page_size INT, @page_index INT; SELECT @page_size="&page_size&", @page_index="&page_index&"; WITH XMLNAMESPACES('http://panax.io/xover' as xo, 'http://panax.io/source' as __data, 'http://panax.io/state' as state, 'http://panax.io/metadata' as meta, 'http://panax.io/custom' as custom, 'http://panax.io/fetch/request' as source, 'http://www.mozilla.org/TransforMiix' as transformiix"&namespaces&" ), #table AS ( SELECT [@meta:position]=ROW_NUMBER() OVER(ORDER BY "&order_by&"), [@meta:resultCount] = COUNT(1) OVER(), * FROM ( SELECT [@meta:totalCount] = COUNT(1) OVER(), "&data_fields&" FROM "&command&" "&data_predicate&") #table "&extra_predicate&") SELECT [@meta:pageIndex]=@page_index, [@meta:pageSize]=@page_size, [@meta:totalCount]=(SELECT TOP 1 [@meta:totalCount] FROM #table), [@meta:resultCount]=(SELECT TOP 1 [@meta:resultCount] FROM #table), ( SELECT * FROM #table "&max_records_predicate&" ORDER BY 1 OFFSET @page_size * (@page_index-1) ROWS FETCH NEXT @page_size ROWS ONLY FOR XML PATH('xo:r'), TYPE) FOR XML PATH('__data:rows'), TYPE"
ELSEIF INSTR(sType,"F")<>0 THEN
    IF INSTR(content_type,"xml")>0 THEN
        strSQL=strSQL&"WITH XMLNAMESPACES('http://panax.io/xover' as xo, 'http://panax.io/state' as state, 'http://panax.io/metadata' as meta, 'http://panax.io/custom' as custom, 'http://panax.io/fetch/request' as source, 'http://www.mozilla.org/TransforMiix' as transformiix) SELECT (SELECT "&command & data_predicate&" FOR XML PATH(''), TYPE) FOR XML PATH('xo:response'), TYPE"
    ELSE
        strSQL="SELECT "&command & data_predicate
    END IF
ELSE
    strSQL="SELECT "&data_fields&" FROM "&command & " AS Result "&data_predicate
END IF

strSQL = REPLACE(strSQL, "'NULL'", "NULL")
strSQL = REPLACE(strSQL, "'null'", "null")
'strSQL = "BEGIN TRY "EXECUTE AS USER='"&session("user_login")&"' END TRY BEGIN CATCH END CATCH; "
strSQL = sParamsDeclaration &"SET NOCOUNT ON; "& sParamsDefinition &strSQL

'strSQL="BEGIN TRY "&strSQL&" END TRY BEGIN CATCH DECLARE @Message NVARCHAR(MAX); SELECT @Message=ERROR_MESSAGE(); EXEC [$Table].[getCustomMessage] @Message=@Message, @Exec=1; END CATCH"
'ELSE
'    IF INSTR(content_type,"xml")>0 THEN
'        IF INSTR(sType,"T")<>0 THEN 'Table  y Table Function
'            strSQL="(SELECT "&data_fields&" FROM "&sRoutineName & " "&data_predicate&" ORDER BY 1 FOR XML PATH('xo:r'), TYPE)"
'        END IF
'        strSQL="SET NOCOUNT ON; WITH XMLNAMESPACES('http://panax.io/xover' as xo, 'http://panax.io/fetch/request' as source, 'http://www.mozilla.org/TransforMiix' as transformiix) SELECT "&strSQL&" FOR XML PATH('')"& root_node &", TYPE"
'    ELSE
'        strSQL="SET NOCOUNT ON; SELECT "&strSQL&" AS Result"
'    END IF
strSQL = replaceMatch(strSQL,"<(DEFAULT|NULL)/>","$1")
IF INSTR(SESSION("user_login"),"@panax.io")<>0 THEN %>
<!--<%= strSQL  %> -->
<% 
END IF
IF 1=0 AND INSTR(SESSION("user_login"),"@panax.io")<>0 AND Debug THEN
    %><!-- <%
    for each x in Request.ServerVariables%>
 <%= "<B>" & x & ":</b> " & Request.ServerVariables(x) & "<p />" %>
<%  next %> -->
<%  response.end
END IF
IF INSTR(content_type,"xml")>0 THEN
    Response.CodePage = 65001
    Response.CharSet = "UTF-8"
    response.ContentType = "text/xml" 
END IF
ON ERROR RESUME NEXT
SET recordset = oCn.Execute(strSQL)
IF Err.Number<>0 THEN 
    IF INSTR(content_type,"xml")>0 THEN
        response.write "<!--"&strSQL&"-->" & vbcrlf
    END IF
    manageError(Err)
    response.end
END IF
DIM r: r=0
DO
    r = r+1
    IF r>1 THEN
        Response.Clear()
    END IF
    IF INSTR(Response.ContentType,"xml")>0 THEN
        IF debug THEN
            response.write "<!--"&recordset.Source&"-->"
        END IF
    END IF
    IF Err.Number<>0 THEN 'OR r>max_recordsets THEN 
        manageError(Err)
    ELSEIF recordset.fields.Count>0 THEN 
        IF NOT (recordset.BOF and recordset.EOF) THEN %>
<%          
            ON ERROR GOTO 0
            DIM oField, sDataType, sValue %>
<%          IF INSTR(content_type,"xml")>0 THEN
                'response.Write("<?xml version='1.0' encoding='UTF-8'?>")
                oField = recordset(0)
                IF NOT(ISNULL(oField)) THEN
                    oXMLFile.LoadXML(oField)
                END IF
                IF oXMLFile.documentElement IS NOTHING THEN
                    IF Request.ServerVariables("HTTP_ROOT_NODE")<>"" THEN %>
<<%= Request.ServerVariables("HTTP_ROOT_NODE") %> xmlns:xo="http://panax.io/xover" xmlns:source="http://panax.io/fetch/request" />
<%                  ELSE
                            Response.Status = "204 No Content"
                    END IF
                END IF
                'oXMLFile.loadXML(oXMLFile.transformNode(xslValues))
                xslFile=server.MapPath(".")&"\normalize_namespaces.xslt"
                IF (fso.FileExists(xslFile)) THEN
                    Set xslDoc=Server.CreateObject("Microsoft.XMLDOM")
                    xslDoc.async="false"
                    xslDoc.load(xslFile)
                    oXMLFile.loadXML(oXMLFile.transformNode(xslDoc))
                END IF
                'response.write "  Cache-Response: "&Request.ServerVariables("Cache-Response")&"-->"
                'response.write "<!-- Cache-Response: "&Request.ServerVariables("HTTP_CACHE_RESPONSE")&"-->"
                DIM x_parameters: set x_parameters = oXMLFile.selectNodes("/xo:parameters/*")
                IF x_parameters.length>0 THEN
                    FOR EACH oNode IN x_parameters
                        ON ERROR RESUME NEXT
                        IF "@"&oNode.nodeName <> output_parameter THEN
                            Response.AddHeader "x-"&replace(oNode.nodeName,"_","-"), oNode.firstChild.xml
                        END IF
                        IF Err.Number<>0 THEN
                            response.write "<!--<WARNING>"& Err.Description &"</WARNING>-->"
                        END IF
                        ON ERROR GOTO 0
		            NEXT
                    FOR EACH oNode IN x_parameters
                        IF "@"&oNode.nodeName = output_parameter OR output_parameter="" AND x_parameters.length = 1 THEN
                            Response.Clear()
                            DIM objFirstChild:	set objFirstChild = Server.CreateObject("Microsoft.XMLDOM")
                            objFirstChild.Async = false
                            objFirstChild.loadXML(oNode.firstChild.xml)
                            ON ERROR RESUME NEXT
                            IF 1=1 or Request.ServerVariables("HTTP_CACHE_RESPONSE")="true" THEN
                                'response.write "<!-- Saved to "&file_location&"-->" & vbcrlf
                                objFirstChild.save file_location
                                objFirstChild.save alt_file_location
                            END IF
                            Response.write objFirstChild.xml
                            IF Err.Number<>0 THEN
                                response.write "<!--<WARNING>"& Err.Description &"</WARNING>-->"
                            END IF
                            ON ERROR GOTO 0
                        END IF
		            NEXT
                ELSE
                    ON ERROR RESUME NEXT
                    IF 1=1 or Request.ServerVariables("HTTP_CACHE_RESPONSE")="true" THEN
                        'response.write "<!-- Saved to "&file_location&"-->" & vbcrlf
                        oXMLFile.save file_location
                        oXMLFile.save alt_file_location
                    END IF
                    IF Err.Number<>0 THEN
                        response.write "<!--<WARNING>"& Err.Description &"</WARNING>-->"
                    END IF
                    response.write oXMLFile.xml
                    ON ERROR GOTO 0
                END IF
            ELSE %>
                [<% dim f: f=0: DO UNTIL recordset.EOF 
                    f = f + 1 %>
                {"#":<%= f %>
<% FOR EACH oField IN recordset.fields 
                        IF oField.name="" THEN 
                        END IF 
                        IF TypeName(oField)="Field" THEN 
                            sDataType=TypeName(oField.value): sValue=oField.value 
                        ELSE 
                            sDataType=TypeName(oField): sValue=oField 
                        END IF %>
		                    , "<%= oField.name %>":<% SELECT CASE UCASE(sDataType): CASE "NULL": %>null<% CASE "BOOLEAN": %><% IF sValue THEN %>true<% ELSE %>false<% END IF %><% CASE ELSE %>"<%= RTRIM(REPLACE(replaceMatch(sValue, "["&chr(13)&""&chr(10)&""&vbcr&""&vbcrlf&"]", ""&vbcrlf),"""", """")) %>"<% END SELECT %> 
	                    <% NEXT %>
                    }
                    <% recordset.MoveNext
 	                LOOP %>]
<% recordset.Close 
             END IF 
        ELSE 
            IF NOT(debug) THEN
                Response.Status = "204 No Content" 
            END IF
        END IF 
    ELSE 
        IF INSTR(content_type,"xml")>0 THEN
            Response.CodePage = 65001
            Response.CharSet = "UTF-8"
            response.ContentType = "text/xml" 
%><?xml-stylesheet type="text/xsl" href="message.xslt" role="message" target="body" action="append" ?>
<xo:message xmlns:xo="http://panax.io/xover" xo:id="message_<%= REPLACE(REPLACE(REPLACE(NOW(),":",""),"/","")," ","_") %>" type="success">El proceso ha terminado</xo:message>
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
ON ERROR GOTO 0
%>