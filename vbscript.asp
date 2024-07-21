<%
Const adInteger = 3
Const adVarChar = 200
Function Base64Encode(sText)
    Dim oXML, oNode

    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.nodeTypedValue =Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function

Function Base64Decode(ByVal vCode)
    Dim oXML, oNode

    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.CreateElement("base64")
    oNode.dataType = "bin.base64"
    oNode.text = vCode
    Base64Decode = Stream_BinaryToString(oNode.nodeTypedValue)
    Set oNode = Nothing
    Set oXML = Nothing
End Function

'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string 
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To text/string
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the output text (unicode) data.
  BinaryStream.CharSet = "us-ascii"

  'Open the stream And get text/string data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function

Function Hash(HashType, Target)
    On Error Resume Next

    Dim PlainText

    If IsArray(Target) = True Then PlainText = Target(0) Else PlainText = Target End If

    With CreateObject("ADODB.Stream")
    .Open
    .CharSet = "Windows-1252"
    .WriteText PlainText
    .Position = 0
    .CharSet = "UTF-8"
    PlainText = .ReadText
    .Close
    End With

	If Err.number<>0 Then
		PlainText = REPLACE(encodeURL(PlainText),"%","")
		Hash HashType, PlainText
	Else
		Set UTF8Encoding = CreateObject("System.Text.UTF8Encoding")
		Dim PlainTextToBytes, BytesToHashedBytes, HashedBytesToHex

		PlainTextToBytes = UTF8Encoding.GetBytes_4(PlainText)

		Select Case HashType
		Case "md5": Set Cryptography = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider") '< 64 (collisions found)
		Case "ripemd160": Set Cryptography = CreateObject("System.Security.Cryptography.RIPEMD160Managed")
		Case "sha1": Set Cryptography = CreateObject("System.Security.Cryptography.SHA1Managed") '< 80 (collision found)
		Case "sha256": Set Cryptography = CreateObject("System.Security.Cryptography.SHA256Managed")
		Case "sha384": Set Cryptography = CreateObject("System.Security.Cryptography.SHA384Managed")
		Case "sha512": Set Cryptography = CreateObject("System.Security.Cryptography.SHA512Managed")
		Case "md5HMAC": Set Cryptography = CreateObject("System.Security.Cryptography.HMACMD5")
		Case "ripemd160HMAC": Set Cryptography = CreateObject("System.Security.Cryptography.HMACRIPEMD160")
		Case "sha1HMAC": Set Cryptography = CreateObject("System.Security.Cryptography.HMACSHA1")
		Case "sha256HMAC": Set Cryptography = CreateObject("System.Security.Cryptography.HMACSHA256")
		Case "sha384HMAC": Set Cryptography = CreateObject("System.Security.Cryptography.HMACSHA384")
		Case "sha512HMAC": Set Cryptography = CreateObject("System.Security.Cryptography.HMACSHA512")
		End Select

		Cryptography.Initialize()

		If IsArray(Target) = True Then Cryptography.Key = UTF8Encoding.GetBytes_4(Target(1))

		BytesToHashedBytes = Cryptography.ComputeHash_2((PlainTextToBytes))

		For x = 1 To LenB(BytesToHashedBytes)
		HashedBytesToHex = HashedBytesToHex & Right("0" & Hex(AscB(MidB(BytesToHashedBytes, x, 1))), 2)
		Next

		If Err.Number <> 0 Then Response.Write(Err.Description) Else Hash = LCase(HashedBytesToHex)

		On Error GoTo 0
	End if

End Function

Sub CreateFolder(ByVal FullPath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(FullPath) Then
        CreateFolder fso.GetParentFolderName(FullPath)
        fso.CreateFolder FullPath
    End If
End Sub

Function transformXML(xmlfile, xslfile)
	Dim xDoc, xslDoc
	Set xDoc = Server.CreateObject("Microsoft.XMLDOM")
	xDoc.async="false"

	IF TypeName(xmlfile)="DOMDocument" THEN
		Set xDoc = xmlfile
	ELSEIF TypeName(xmlfile)="IXMLDOMElement" THEN
		xDoc.loadXML(xmlfile.xml)
    ELSE
		'Load XML file
		xDoc.async = false
		xDoc.load(xmlfile)
	END IF
	'Load XSL file
	set xslDoc = Server.CreateObject("Microsoft.XMLDOM")
	xslDoc.async = false
	xslDoc.load(xslfile)
	'Transform file
	transformXML = xmlDoc.transformNode(xslDoc)
End function

function asyncCall(strUrl)
    Set xmlHttp = Server.Createobject("MSXML2.ServerXMLHTTP")
    xmlHttp.Open "GET", strUrl, False
    xmlHttp.setRequestHeader "User-Agent", "asp httprequest"
    xmlHttp.setRequestHeader "content-type", "application/x-www-form-urlencoded"
    xmlHttp.Send
    'response.write xmlHttp.responseText
    'xmlHttp.abort()
    set xmlHttp = Nothing   
end function 

FUNCTION checkConnection(oCn)
	IF LCASE(SESSION("secret_engine")) = "google" THEN
		DIM google_response
		google_response = apiCall("https://www.googleapis.com/oauth2/v3/tokeninfo?id_token=" & session("secret_token"))
		Set xml_google_response = JSONToXML(google_response)
		IF NOT (xml_google_response.documentElement.selectSingleNode("error_description") IS NOTHING) OR NOT(xml_google_response.documentElement.selectSingleNode("email").text = session("user_login")) THEN
			session.Abandon
		END IF
		IF session("user_login") = "" THEN
			Session("AccessGranted") = FALSE
			session("status") = "unauthorized"
	        Response.ContentType = "application/json"
			Response.CharSet = "ISO-8859-1"
			Response.Status = "401 Unauthorized" %>
			{
			"message": "Conexión no autorizada"
			}
		<% 	response.end
		END IF
	ELSE
		DIM StrCnn: StrCnn = "driver={SQL Server};server="&SESSION("secret_server_id")&";uid="&SESSION("secret_database_user")&";pwd="&SESSION("secret_database_password")&";database="&SESSION("secret_database_name")
		If oCn.State = 0 THEN
			ON ERROR RESUME NEXT
			oCn.Open StrCnn
		End if
		IF NOT(Err.Number=0 AND (TRIM(SESSION("secret_server_id"))<>"" AND TRIM(SESSION("secret_database_user"))<>"" AND TRIM(SESSION("secret_database_password"))<>"" AND TRIM(SESSION("secret_database_name"))<>"")) THEN 
			Session("AccessGranted") = FALSE
			Session("status") = "unauthorized"
			DIM error_description
			IF (Err.number<>0) THEN
				error_description = Err.Description
			ELSEIF oCn.errors.count<>0 THEN
				error_description = "Can't connect"
			END IF
			ErrorDesc=SqlRegEx.Replace(error_description, "")
			'response.write Err.Number&": "&Err.Description
			IF INSTR(ErrorDesc,"SQL Server does not exist or access denied")>0 OR INSTR(ErrorDesc,"Communication link failure")>0 OR INSTR(ErrorDesc,"ConnectionWrite")>0 THEN
				AsyncCall "http://localhost:8080/startSQL"
				'AsyncCall Left(currentLocation, instrRev(currentLocation, "/"))&"reconnect.asp"
				Err.Clear
				Sleep(3)
				If oCn.State = 0 Then
					'response.write "Here 1 "&oCn.State&"<br/>"
					ON ERROR RESUME NEXT
					oCn.Open StrCnn
					IF Err.Number<>0 THEN 
						'response.write "Here 2 "&oCn.State
						Response.ContentType = "application/json"
						Response.CharSet = "ISO-8859-1"
						ErrorDesc=SqlRegEx.Replace(Err.Description, "")
						'response.Write ErrorDesc
						IF INSTR(ErrorDesc,"SQL Server does not exist or access denied")>0 OR INSTR(ErrorDesc,"Communication link failure")>0 THEN
							Response.Status = "503 Service Unavailable" '"408 Request Timeout"
	%>
						{
						"success": false,
						"message": "No se pudo establecer una conexión con la base de datos <%= sDatabaseName %>: <%= RegEx_JS_Escape.Replace(SqlRegEx.Replace(Err.Description, ""), "\$&") %>"
						}
					<% 	response.end
						END IF
					END IF
				End If
			END IF
		END IF
	END IF
END FUNCTION


Function RandomNumber(intHighestNumber)
	Randomize
	RandomNumber = Int(Rnd * intHighestNumber) + 1
End Function

Function testMatch(sOriginal, sPatrn)
	Dim regEx, Match, Matches, strReturn
	Set regEx = New RegExp
	regEx.Pattern = sPatrn
	regEx.IgnoreCase = True				' Distinguir mayúsculas de minúsculas.
	regEx.Multiline = True				' Distinguir mayúsculas de minúsculas.
	regEx.Global = True
	testMatch = regEx.Test(sOriginal)
End Function

Function getMatch(sOriginal, sPatrn)
	Dim regEx, Match, Matches, strReturn
	Set regEx = New RegExp
	regEx.Pattern = sPatrn
	regEx.IgnoreCase = True				' Distinguir mayúsculas de minúsculas.
	regEx.Multiline = True				' Distinguir mayúsculas de minúsculas.
	regEx.Global = True
	Set Matches = regEx.Execute(sOriginal) 
	Set getMatch = Matches
End Function

Function replaceMatch(sOriginal, sPatrn, sReplacementText)
	Dim regEx, Match, Matches, strReturn
	Set regEx = New RegExp
	regEx.Pattern = sPatrn
	regEx.IgnoreCase = True				' Distinguir mayúsculas de minúsculas.
	regEx.Multiline = True				' Distinguir mayúsculas de minúsculas.
	regEx.Global = True
	IF IsNullOrEmpty(sOriginal) THEN
		replaceMatch = ""
	ELSE
		replaceMatch = regEx.Replace(sOriginal, sReplacementText) 
	END IF
End Function

Function replaceEvaluatingMatch(sOriginal, sPatrn, sReplacementText)
	Dim regEx, Match, Matches, strReturn
	Set regEx = New RegExp
	regEx.Pattern = sPatrn
	regEx.IgnoreCase = True				' Distinguir mayúsculas de minúsculas.
	regEx.Multiline = True				' Distinguir mayúsculas de minúsculas.
	regEx.Global = True
	IF IsNullOrEmpty(sOriginal) THEN
		replaceEvaluatingMatch = ""
	ELSE
		replaceEvaluatingMatch = regEx.Replace(sOriginal, sReplacementText) 'Evaluate(regEx.Replace(sOriginal, sReplacementText))
	END IF
End Function

Function applyTemplate(sOriginal, sPatrn, sTemplate)
	Dim regEx, Match, Matches, strReturn
	Set regEx = New RegExp
	regEx.Pattern = sPatrn
	regEx.IgnoreCase = True				' Distinguir mayúsculas de minúsculas.
	regEx.Multiline = True				' Distinguir mayúsculas de minúsculas.
	regEx.Global = True
	strReturn=regEx.Replace(sOriginal, sTemplate)
	strReturn=REPLACE(strReturn, "ñ", "ni")
	strReturn=REPLACE(strReturn, "Ñ", "NI")
	applyTemplate=EVAL(strReturn)
End Function

Function getDisplayName(strTemp)
	Dim patrn
	patrn="\{(.*)\}*"
	Dim regEx, Match, Matches
	Set regEx = New RegExp
	regEx.Pattern = patrn
	regEx.IgnoreCase = True				' Distinguir mayúsculas de minúsculas.
	regEx.Multiline = True				' Distinguir mayúsculas de minúsculas.
	regEx.Global = True
	strTemp=regEx.Replace(strTemp, "")
	getDisplayName=strTemp
End Function

Function getParameters(ByVal sParameters)
	Set getParameters = getMatch(sParameters, "(?:@)([\w\.\(\)""]*)=([^@]*)(?!(\s*@))")
'	Set getParameters = getMatch(sParameters, "(?:@)([\w\.\(\)""]*)=([\w\d\s,=\(\)\[\]""'\.\*\+\/\-\\&\?\!<>]*)(?!(\s*@))")
End Function

Function getMetadataString(sOriginal, sColumnName)
	Dim patrn, regEx, Match, Submatch, sMetadataString
	patrn=","&sColumnName&"@{(.*?)}@"
'	strMatchPattern="«\w*(\([^«]|[\w\(\)\,\s\-]*\))*»"
'	Debugger Me, "<strong>"&sColumnName&"("&getMatch(","&sOriginal, patrn).Count&"): </strong> ("&sOriginal&"): "
'	response.write sOriginal &"<br>"
	For Each Match in getMatch(","&sOriginal, patrn)
		FOR EACH Submatch IN Match.Submatches
			getMetadataString=TRIM(Submatch)
		NEXT
	Next 
End Function

Function getMetadata(sOriginal, sColumnName, sPropertyName)
'	sOriginal=getMetadataString(sOriginal, sColumnName)
	Dim strMatchPattern, i, Match, sValue
'	response.write sOriginal & "<br>"
	'IF sPropertyName="ControlParameters" THEN Debugger Me, sOriginal
	strMatchPattern=";@"&sPropertyName&"\:(.*?);@"
	'IF sPropertyName="ControlParameters" THEN Debugger Me, strMatchPattern
'	strMatchPattern="«\w*(\([^«]|[\w\(\)\,\s\-]*\))*»"
	DIM SubMatch
	i=0
	For Each Match in getMatch(";"&sOriginal&"@", strMatchPattern)
		'For Each SubMatch IN Match.SubMatches
			sValue=TRIM(Match.SubMatches(0))
		'NEXT
'		i=i+1
'		sValue=LEFT(Match.value, LEN(Match.value)-1)
'		sValue=RTRIM(replace(sValue, ";"&sPropertyName&":", ""))
	Next 
'	Debugger Me, sPropertyName&"> "&sValue
'	response.write "<br><br>"
	IF sValue="NULL" THEN sValue=NULL
	getMetadata=sValue
End Function

Function Evaluate(ByVal sInput)
	Evaluate=fncEvaluate(sInput)
End Function
Function fncEvaluate(ByVal sInput)
	DIM vReturnValue
'	IF IsObject(vInput) THEN
'		EXECUTE("Set Evaluate=sInput")
'	ELSE
	ON ERROR  RESUME NEXT
	EXECUTE("vReturnValue="&CString(sInput).RemoveEntities())
	IF Err.Number<>0 THEN
		response.write "Ocurrió el siguiente error en funcion <strong>fncEvaluate</strong>:"&Err.Description&vbcrlf&"<br> Al evaluar "&sInput&".<br>"
		Debugger Me, ("vReturnValue="&CString(sInput).RemoveEntities().Replace("(["&chr(13)&""&chr(9)&""&chr(10)&""&vbcr&""&vbcrlf&""&vbtab&"])", "<strong>-especial-</strong>"))
		response.end
		Err.Clear
	END IF
	ON ERROR  GOTO 0
'	END IF
	fncEvaluate=vReturnValue
End Function


function evalTemplate(byVal fldformat, ByRef oDictionary, ByVal aDataRow)
	Dim patrn, fldvalue
	patrn="\{(\w*)\}*"
	Dim regEx, Match, Matches
	Set regEx = New RegExp
	regEx.Pattern = patrn
	regEx.IgnoreCase = True				' Distinguir mayúsculas de minúsculas.
	regEx.Multiline = True				' Distinguir mayúsculas de minúsculas.
	regEx.Global = True
''				response.write fldformat &"-->"
'	fldvalue=regEx.Replace(fldformat, "getDataRowValue(RowNumber, oDictionary(""fieldsDictionary"")(""$1""))")
	fldvalue=regEx.Replace(fldformat, "getDataRowValue(aDataRow, oDictionary.item(""$1"").ParentCell.ColumnNumber)")
''                fldvalue=replace(fldformat, "{0}", fldvalue) '"datarow(oDictionary(""fieldsDictionary"")(""PrecioViv""))-123")
	evalTemplate=EVAL(fldvalue)
End Function

function EvaluateTemplate(byVal fldformat, ByRef oDictionary, ByVal iRecord)
	Dim patrn, fldvalue
	patrn="\{(\w*)\}*"
	Dim regEx, Match, Matches
	Set regEx = New RegExp
	regEx.Pattern = patrn
	regEx.IgnoreCase = True				' Distinguir mayúsculas de minúsculas.
	regEx.Multiline = True				' Distinguir mayúsculas de minúsculas.
	regEx.Global = True
'	Dim a, i
'	a=oDictionary.Keys
'	for i=0 to oDictionary.Count-1
'	  Response.Write("a: "&a(i))
'	  Response.Write("<br />")
'	next
'	set a=nothing
	fldvalue=regEx.Replace(fldformat, "oDictionary.item(""$1"")")'oDictionary.item(""$1"").Value")
'	response.write fldvalue&"<br>"
	EvaluateTemplate=EVAL(fldvalue)
End Function

Function TranslateTemplate(ByVal sTemplate)
	TranslateTemplate=CString(sTemplate).Remove(vbcrlf).Replace("\""", """""").Replace("\$_GET\[""(.*?)""\]", "request.querystring(""$1"")").Replace("([^\\])\[(?!\$)(\w*)\](?!=\!)", "$1oFields(""$2"")").RemoveEntities().Replace("((?=.?)(?!\\)\[(?!\$)(?:\w*)(?!\\)\])(\!)", "$1") '\[(?!\$)(\w*)\]
End Function

function EvaluateFieldsTemplate(ByVal sTemplate, ByRef oFields)
	ON ERROR  RESUME NEXT: 
	IF sTemplate<>"" THEN
		Assign EvaluateFieldsTemplate, EVAL(sTemplate)
		'EXECUTE("SET EvaluateFieldsTemplate."&sPropertyName&"=Evaluate(sTemplate)")
		'IF Err.Number<>0 THEN
		'	Err.Clear
		'	EXECUTE("EvaluateFieldsTemplate."&sPropertyName&"=sTemplate")
		'END IF
	ELSE
		EvaluateFieldsTemplate=sTemplate
	END IF
	[&Catch] TRUE, Me, "EvaluateFieldsTemplate", ""&sTemplate&" <strong>=></strong><br> ": ON ERROR  GOTO 0
End Function 

function ContextualEvaluation(ByVal sTemplate, ByRef oContext)
	ON ERROR  RESUME NEXT: 
	IF sTemplate<>"" THEN
		Assign EvaluateFieldsTemplate, Evaluate(sTemplate)
		EXECUTE("SET EvaluateFieldsTemplate."&sPropertyName&"=Evaluate(sTemplate)")
		IF Err.Number<>0 THEN
			Err.Clear
			EXECUTE("oElement."&sPropertyName&"=vInput")
		END IF

	ELSE
		EvaluateFieldsTemplate=sTemplate
	END IF
	[&Catch] TRUE, Me, "EvaluateFieldsTemplate", ""&sTemplate&" <strong>=></strong><br> ": ON ERROR  GOTO 0
End Function 

function EvaluateRowTemplate(byVal sTemplate, ByRef oFields)
	Dim fldvalue
	fldvalue=TranslateTemplate(sTemplate)
'	fldvalue=regEx.Replace(sTemplate, "oFields(""$1"")")'oDictionary.item(""$1"").Value")

'	response.write fldvalue&"<br>"
	ON ERROR  RESUME NEXT: 
	DIM oResult
	Assign oResult, EVAL(fldvalue)
	[&Catch] TRUE, Me, "EvaluateRowTemplate", ""&sTemplate&" <strong>=></strong><br> "&fldvalue: ON ERROR  GOTO 0

	IF IsObject(oResult) THEN
		Set EvaluateRowTemplate=oResult
	ELSE
		EvaluateRowTemplate=oResult
	END IF
End Function

Function TextDataBind(byRef sText, ByRef oFields)

'"& (""&RTRIM( Candidato )&"") &" 
'RESPONSE.WRITE  eval("""Texto: ""&Evaluate(""RTRIM( ""& """""""&Candidato&""""""" &"" )"")&"" Fin Texto""")
'"&RTRIM(""«Candidato»""))&"



'RESPONSE.WRITE  EVAL("""Texto: ""&RTRIM(Candidato)&"" Fin Texto""")
'RESPONSE.WRITE  EVAL("""Texto: ""&Evaluate("""&REPLACE("""&RTRIM(""""&Candidato&"""")&""", """", """""""")&""")&"" Fin Texto""")
'RESPONSE.WRITE Evaluate("""Nombre: ""& (""&( Candidato )&"") &""""" )
'RESPONSE.END

'	Set TextDataBind = CString(sText).Replace("""", """""").Append("""").Prepend("""").Replace("(?:&lt;|<)%(.*?)%(?:&gt;|>)", """&Evaluate(""($1)"") &""" ).Replace("(?:&laquo;|«)(.*?)(?:&raquo;|»)", """&(""$1"")&""").Evaluate()
'"""& (""Evaluate( $1 ) "") &"""
	IF NOT oFields IS NOTHING THEN
		Set TextDataBind = CString(sText).DoubleQuote() _
			.Replace("""(?:&laquo;|«)(.*?)(?=[\s]*(?:&raquo;|»))", """") _
			.Replace("\[(.*?)\](?=[\)\s]*(?=&raquo;|»))", "oFields(""$1"").GetCode()") _
			.Replace("\{(.*?)\}(?=[\)\s]*(&raquo;|»))", "oFields(""$1"")") _
			.Replace("#(.*?)#(?=&raquo;|»)", "session(""$1"")") _
			.Replace("(?:&laquo;|«)(.*?)(?:&raquo;|»)", """&( $1 )&""") _
			.Evaluate()
	ELSE
		Set TextDataBind = CString(sText).DoubleQuote() _
			.Replace("(?:&laquo;|«)(.*?)(?:&raquo;|»)", """&( $1 )&""") _
			.Evaluate()
	END IF
End Function

Function TextEvaluate(byRef sText)
	Set TextEvaluate = CString(sText).Replace("""", """""").Append("""").Prepend("""").Replace("(?:&laquo;|«)(.*?)(?:&raquo;|»)", """&($1)&""").Evaluate()
End Function

Function getDataRowValue(ByVal aDataRow, ByVal FieldColumnNumber)
	Dim thisrow, returnValue
	IF FieldColumnNumber="" THEN
		returnValue=""
	ELSE
'        thisrow=arrayData(RowNumber)
'		response.write "RowNumber: "&RowNumber&", FieldColumnNumber: "&FieldColumnNumber&" ("&thisrow&")<br>"
'		aDataRow=split(thisrow, "#c#") 
		returnValue=aDataRow(FieldColumnNumber)
	END IF
	IF Err.Number<>0 THEN
		response.write "Error en getDataRowValue("&RowNumber&", "&FieldColumnNumber&") <br>"
		'Err.Clear
	END IF

	getDataRowValue=returnValue
End Function

Function calculateRowSpan(ByVal RowNumber, ByVal spanRowsBy)
'	response.write "Buscando a partir de n: "&RowNumber&"<br>"
	returnValue=0
	IF spanRowsBy="" THEN
		calculateRowSpan=1
			EXIT FUNCTION
	ELSE
	    FOR therow=RowNumber TO ubound(arrayData)-1
			thisrow=arrayData(therow)
	        datarow=split(thisrow, "#c#")
			IF NOT(therow=RowNumber) THEN
				thisRowValReference=evalTemplate(spanRowsBy, therow)
'				response.write "r: "&therow&", n: "&RowNumber&", "&UBOUND(arrayData)&"("&thisRowValReference&" vs "&lastRowValReference&"): "&returnValue&"<br>"
				IF NOT(lastRowValReference=thisRowValReference) THEN
'					response.write "Valor encontrado: "&returnValue
					calculateRowSpan=returnValue
					EXIT FUNCTION
				END IF
			END IF
			lastRowValReference=evalTemplate(spanRowsBy, therow)
			returnValue=returnValue+1
		NEXT
		calculateRowSpan=returnValue
	END IF
End Function


FUNCTION ErrorDisplay(parmSource, parmConn, objDictionary)
'    If objDictionary.item("debug")=true THEN    
'        response.write "<hr>ErrorDisplay called<br>"
'        response.flush
'    END IF
    ErrorDisplay=0
    DIM errvbs, errdesc
    errvbs=err.number
    errdesc=err.description
    objDictionary.item("errorsource")=parmSource
'     IF objDictionary.item("debug")=true THEN    
'        response.write "errvbs=" & errvbs & "<br>"
'        response.write "errdesc=" & errdesc & "<br>"
'        response.write "parmsource=" & parmSource & "<br>"
'        response.flush
'    END IF    
    DIM errordetails, customerror
    customerror=false
    If errvbs<>0 THEN
        SELECT CASE errvbs
            CASE -2147467259
                objDictionary.item("errordesc")="Bad DSN" 
                objDictionary.item("errornum")=2
                errordetails=objDictionary.item("conn")
                ErrorDisplay=2
                objDictionary.item("errorname")="error_dsn_bad"
            CASE -2147217843
                objDictionary.item("errordesc")="Bad DSN Login Info"  
                objDictionary.item("errornum")=3
                errordetails=objDictionary.item("conn")
                ErrorDisplay=3
                objDictionary.item("errorname")="error_dsn_bad_login"
            CASE -2147217865
                objDictionary.item("errordesc")="Invalid Object Name"
                objDictionary.item("errornum")=4
                errordetails="probably query has wrong table name - SQL= " & objDictionary.item("sql")
                ErrorDisplay=4
                objDictionary.item("errorname")="error_query_badname"
            CASE -2147217900
              objDictionary.item("errordesc")="Bad Query Syntax"
                objDictionary.item("errornum")=5
                errordetails=errdesc & " - SQL= " & objDictionary.item("sql")
                ErrorDisplay=5
                objDictionary.item("errorname")="error_query_badsyntax"
            CASE ELSE
                objDictionary.item("errordesc")="VBscript Error #=<b>" & errvbs & "</b>, desc=<b>" & errdesc & "</b>"
                errordetails="n/a"
                ErrorDisplay=1
                objDictionary.item("errorname")="error_unexpected"
        END SELECT
    END IF
'     IF objDictionary.item("debug")=true THEN    
'           response.write "objDictionary.item(""errordesc"")=" & objDictionary.item("errordesc") & "<br>"
'            response.write "objDictionary.item(""errornum"")=" & objDictionary.item("errornum") & "<br>"
'            response.write "errordetails=" & errordetails & "<br>"
'            response.write "errorDisplay=" & errordisplay & "<br>"
'            response.write "objDictionary.item(""errordesc"")=" & objDictionary.item("errordesc") & "<br>"
'    END IF    
    
    Dim errorname
    errorname=objDictionary.item("errorname")
    IF objDictionary.item(errorname)="" THEN
        ' nothing to do
    ELSE
        customerror=true
        objDictionary.item("errordesc")=objDictionary.item(errorname)
    END IF

    IF customerror=TRUE THEN
            objDictionary.item("errordesc")=    replace(objDictionary.item("errordesc"), "{details}", errordetails)
    ELSE
        IF objDictionary.item("errorsdetailed")=TRUE THEN    
            objDictionary.item("errordesc")=objDictionary.item("errordesc") & " details=<b>" & errordetails & "</b>"
        END IF    
    END IF

    DIM howmanyerrors, dberrnum, dberrdesc, dberrdetails, counter
    howmanyerrors=parmConn.errors.count
'     IF objDictionary.item("debug")=true THEN    
'        response.write "howmanyerrors =" & howmanyerrors & "<br>"
'        response.flush
'    END IF    
    dberrdetails="<b>(details: "

    IF howmanyerrors>0 THEN
        FOR counter= 0 TO 0'howmanyerrors
            dberrnum=parmconn.errors(counter).number
            dberrdesc=parmconn.errors(counter).description
            dberrdetails=dberrdetails & " #=" & dberrnum & ", desc=" & dberrdesc & "; " 
       NEXT
       objDictionary.item("adoerrornum")=1
       objDictionary.item("adoerrordesc")="DB Error " & dberrdetails
    END IF
    'objDictionary.item("errornum")=ErrorDisplay    
'     If objDictionary.item("debug")=true THEN    
'        response.write "objDictionary(""adoerrornum"")=" & objDictionary("adoerrornum") & "<br>"
'        response.write "objDictionary(""adoerrordesc"")=" & objDictionary("adoerrordisc") & "<br>"
'        response.write "Leaving ErrorDisplay Function<br>"
'        response.write "objDictionary(""errornum"")=" & objDictionary("errornum") & "<br>"
'        response.flush
'    END IF
END FUNCTION

Function ToTitleFromPascal(ByVal s)
    Dim s0, s1, s2, s3, s4, sf, Regex
	Set Regex = New RegExp
	Regex.Global = True 
	Regex.IgnoreCase = False 
	regEx.Multiline = True				' Distinguir mayúsculas de minúsculas.
    ' remove name space
	Regex.Pattern = "(.*\.)(.*)"
	s0 = Regex.Replace(s, "$2")

    ' add space before Capital letter
	Regex.Pattern = "[A-Z]"
    s1 = Regex.Replace(s0, " $&")
    
    ' replace '_' with space
	Regex.Pattern = "[_]"
    s2 = Regex.Replace(s1, " ")
    
    ' replace double space with single space
	Regex.Pattern = " "
    s3 = Regex.Replace(s2, " ")
    
    ' remove and double capitals with inserted space
	Regex.Pattern = "([A-Z])\s([A-Z])"
'    response.write s&": "&Regex.Test(s3) &"<br>"
	DO WHILE Regex.Test(s3)
	    s3 = Regex.Replace(s3, "$1$2")
	LOOP
	S4=s3
'    response.write s&": "&Regex.Test(s3) &"<br>" &"<br>"

	Regex.Pattern = "^\s"
    sf = Regex.Replace(s4, "")
    
    ' force first character to upper case
    ToTitleFromPascal=ToTitleCase(sf)
End Function

Function ToTitleCase(ByVal text)
'	RegEx.Replace(RegEx.Replace(@str, "[a-z](?=[A-Z])", "$& ", 0), "(?<=[A-Z])[A-Z](?=[a-z])", " $&", 0)
    Dim sb, i
    
    For i = 0 To LEN(text) - 1
        If i > 0 Then
            If MID(text, i, 1) = " " OR MID(text, i, 1) = vbTab OR MID(text, i, 1) = "/" Then
                sb=sb&(UCASE(MID(text, i+1, 1)))
            Else
                sb=sb&(LCASE(MID(text, i+1, 1)))
            End If
        Else
			sb=sb&UCASE(MID(text, i+1, 1))
        End If
    Next
    
    ToTitleCase=sb
End Function

Function FormatearNombre(strTemp)
	strTemp=replace(UCASE(strTemp), "Á", "A")
	strTemp=replace(UCASE(strTemp), "A", "[AÁ]")
	strTemp=replace(UCASE(strTemp), "É", "E")
	strTemp=replace(UCASE(strTemp), "E", "[EÉ]")
	strTemp=replace(UCASE(strTemp), "Í", "I")
	strTemp=replace(UCASE(strTemp), "I", "[IÍ]")
	strTemp=replace(UCASE(strTemp), "Ó", "O")
	strTemp=replace(UCASE(strTemp), "O", "[OÓ]")
	strTemp=replace(UCASE(strTemp), "Ú", "U")
	strTemp=replace(UCASE(strTemp), "U", "[UÚ]")
	FormatearNombre=strTemp
End Function

'Function FormatValue(ByVal vValue, ByVal sFormat, ByVal iDecimalPositions)
'	IF IsNullOrEmpty(vValue) THEN FormatValue="": Exit Function END IF
'	IF IsNullOrEmpty(sFormat) THEN FormatValue=vValue: Exit Function END IF
'	SELECT CASE UCASE(sFormat)
'	CASE "MONEY"
'		IF IsNullOrEmpty(iDecimalPositions) THEN iDecimalPositions=2
'		FormatValue=FormatCurrency(vValue, iDecimalPositions)
'	CASE "PERCENT"
'		IF IsNullOrEmpty(iDecimalPositions) THEN iDecimalPositions=2
'		FormatValue=FormatPercent(vValue/100, iDecimalPositions)
'	CASE "DATE"
'		FormatValue=FormatDateTime(vValue, 2)
'	CASE "DATETIME"
'		FormatValue=FormatDateTime(vValue, 2)&" "&FormatDateTime(vValue, 3)
'	CASE "NUMERIC"
'		IF IsNullOrEmpty(iDecimalPositions) THEN iDecimalPositions=0
'		FormatValue=FormatNumber(vValue, iDecimalPositions)
'	CASE ELSE
''		IF IsNumeric(vValue) THEN
''			IF IsNullOrEmpty(iDecimalPositions) THEN iDecimalPositions=0
''			FormatValue=FormatNumber(vValue, iDecimalPositions)
''		ELSE
'		FormatValue=vValue
''		END IF
'	END SELECT
'End Function

Function URLDecode2(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
	IF sConvert="" THEN 
       URLDecode2 = ""
       Exit Function
    End If
    If IsNull(sConvert) Then
       URLDecode2 = ""
       Exit Function
    End If
	
    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")
    sOutput = REPLACE(sOutput, "%A0", " ")
    sOutput = REPLACE(sOutput, "%2C", ",")
	
    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")
	
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        'response.write "--"&Left(aSplit(i + 1), 2)&Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2) &"--"
    'response.end
        IF testMatch(Left(aSplit(i + 1), 2),"[0-9A-F]{2}") THEN
            sOutput = sOutput & _
              Chr("&H" & Left(aSplit(i + 1), 2)) &_
              Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
        ELSE 
            sOutput = sOutput & "%" & aSplit(i + 1)
        END IF
      Next
    End If
    URLDecode2 = sOutput
End Function

Function RegExTest(str, patrn)
	Dim regEx
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Pattern = patrn
	RegExTest = regEx.Test(str)
End Function

Function URLDecode(sStr)
	'UrlDecode = URLDecode2(sStr): Exit Function
	DIM str, code, a0
	str=""
	code=sStr&""
	code=Replace(code,"+"," ")
	While len(code)>0
		If InStr(code,"%")>0 Then
			str = str & Mid(code,1,InStr(code,"%")-1)
			code = Mid(code,InStr(code,"%"))
			a0 = UCase(Mid(code,2,1))
			If a0="U" And RegExTest(code,"^%u[0-9A-F]{4}") Then
				str = str & ChrW((Int("&H" & Mid(code,3,4))))
				code = Mid(code,7)
			ElseIf a0="E" And RegExTest(code,"^(%[0-9A-F]{2}){3}") Then
				str = str & ChrW((Int("&H" & Mid(code,2,2)) And 15) * 4096 + (Int("&H" & Mid(code,5,2)) And 63) * 64 + (Int("&H" & Mid(code,8,2)) And 63))
				code = Mid(code,10)
			ElseIf a0>="C" And a0<="D" And RegExTest(code,"^(%[0-9A-F]{2}){2}") Then
				str = str & ChrW((Int("&H" & Mid(code,2,2)) And 3) * 64 + (Int("&H" & Mid(code,5,2)) And 63))
				code = Mid(code,7)
			ElseIf (a0<="B" Or a0="F") And RegExTest(code,"^%[0-9A-F]{2}") Then
				str = str & Chr(Int("&H" & Mid(code,2,2)))
				code = Mid(code,4)
			Else
				str = str & "%"
				code = Mid(code,2)
			End If
		Else
			str = str & code
			code = ""
		End If
	Wend
	URLDecode = str
End Function

Function FormatValue(sParamValue)
	bParameterString=NOT(sParamValue="" OR sParamValue="NULL" OR sParamValue="DEFAULT" OR ISNUMERIC(sParamValue) OR testMatch(sParamValue, "^['@]"))
	IF bParameterString THEN sParamValue="'"&REPLACE(sParamValue,"'","''")&"'" END IF
	IF sParamValue="" THEN sParamValue="NULL" END IF
	FormatValue = sParamValue
End Function

Function encodeURL(sConvert)
	Dim strTemp
	strTemp=server.urlEncode(sConvert)
	strTemp=replace(strTemp, "%2C", ",")
	strTemp=replace(strTemp, "%28", "(")
	strTemp=replace(strTemp, "%29", ")")
	strTemp=replace(strTemp, "%2A", "*")
	strTemp=replace(strTemp, "%3D", "=")
	strTemp=replace(strTemp, "%2E", ".")
	strTemp=replace(strTemp, "%2F", "/")
	strTemp=replace(strTemp, "%3C", "<")
	strTemp=replace(strTemp, "%3E", ">")
	strTemp=replace(strTemp, "%5F", "_")
	encodeURL=strTemp
End Function

Function HTMLEncode(sText)
	HTMLEncode=Server.HTMLEncode(sText)
End Function

Function HTMLDecode(sText)
    Dim i
    sText = Replace(sText, "&quot;", Chr(34))
    sText = Replace(sText, "&lt;"  , Chr(60))
    sText = Replace(sText, "&gt;"  , Chr(62))
    sText = Replace(sText, "&amp;" , Chr(38))
'    sText = Replace(sText, "&nbsp;", Chr(32))
	sText = Replace(sText, "&#x0D;", Chr(13))
    For i = 1 to 255
        sText = Replace(sText, "&#" & i & ";", Chr(i))
    Next
    HTMLDecode = sText
End Function

'ESTAS FUNCIONES DEBEN ESTAR IGUALES QUE EN EL SQL SERVER
Function anioReal(byVal Fecha)
semana=semanaReal(Fecha)
	IF YEAR(Fecha-DATEPART("w", Fecha, 2, 1)+1)<>YEAR(Fecha+7-DATEPART("w", Fecha, 2, 1)) THEN 
		IF semana=1 THEN
			anioReal=YEAR(Fecha+7-DATEPART("w", Fecha, 2, 1))
		ELSE
			anioReal=YEAR(Fecha-DATEPART("w", Fecha, 2, 1)+1)
		END IF
	ELSE
		anioReal=YEAR(Fecha)
	END IF
End Function

Function semanaReal(byVal Fecha)
	Dim iSemanaCalculada
	Fecha=CDATE(Fecha)
	IF DatePart("ww", Fecha, 2, 1)=54 OR DATEPART("w", CDATE("1/1/"&YEAR(Fecha)), 2, 1)<4 THEN 
		iSemanaCalculada=DATEPART("ww", Fecha+7-DATEPART("w", Fecha, 2, 1), 2, 1)
	ELSEIF DATEPART("w", CDATE("1/1/"&YEAR(Fecha)), 2, 1)>4 THEN 
		IF DATEPART("ww", Fecha, 2, 1)=1 AND DATEPART("w", Fecha, 2, 1)>4 THEN 
			iSemanaCalculada=semanaReal(Fecha-DATEPART("w", Fecha, 2, 1)+1)
		ELSE 
			iSemanaCalculada=DATEPART("ww", Fecha-DATEPART("w", Fecha, 2, 1), 2, 1) 
		END IF
	ELSEIF DATEPART("w", CDATE("1/1/"&YEAR(Fecha)), 2, 1)=4 OR DATEPART("w", CDATE("31/12/"&YEAR(Fecha)), 2, 1)=4 THEN 
		iSemanaCalculada=DatePart("ww", Fecha, 2, 1)
	ELSE 
		iSemanaCalculada=-1 
	END IF
	semanaReal=iSemanaCalculada
End Function

Function Numero_Letras(byVal cant)
IF TRIM(cant)="" THEN
	Numero_Letras=""
ELSE
	Numero_Letras=Pesos(CDBL(cant))
END IF
End Function

Function leeArchivo(byVal fileName)
	Const ParaLeer = 1
	Dim fso, f
	Set fso = CreateObject("Scripting.FileSystemObject")
	'ON ERROR  RESUME NEXT
	Set f = fso.OpenTextFile(fileName, ParaLeer)
	IF Err.Number<>0 THEN
		response.write "Error al abrir archivo "&fileName&". Error:"&REPLACE(Err.Description,"'","\'")
		response.end
		''Err.Clear
	END IF
	leeArchivo =  f.ReadAll
	'ON ERROR  GOTO 0
End Function

Sub BindFile (byRef strFile, byVal oRecordSet)
	Dim strMatchPattern, i
	strMatchPattern="«\w*(\([^«]|[\w\(\)\,\s\-]*\))*»"
End Sub

Function interpretaContratos (byRef sContrato) 
	Dim strMatchPattern, i
	Dim Matches, Match
'	sContrato=HTMLDecode(sContrato)
	strMatchPattern="(?:&laquo;|«)(.*?)(?:&raquo;|»)"
	Set Matches = getMatch(sContrato, strMatchPattern)

	i=0
	For Each Match in Matches 
		i=i+1
	'	strReturnStr = i&".- Match found at position " 
	'	strReturnStr = strReturnStr & Match.FirstIndex & ". Match Value is '" 
	'	strReturnStr = strReturnStr & replace(replace(Match.value, "«", ""), "»", "") & "'="&EVAL(replace(replace(Match.value, "«", ""), "»", ""))&"." 
if session("IdUsuario")=1 THEN
'ON ERROR  RESUME NEXT
END IF
		sContrato=replace(sContrato, Match.value, EVAL(Match.Submatches(0)))
	'	sContrato=replace(sContrato, Match.value, "<label style=""text-decoration:'underline';"">&nbsp;&nbsp;&nbsp;"&EVAL(replace(replace(Match.value, "«", ""), "»", ""))&"&nbsp;&nbsp;&nbsp;</label>")
	'	Response.Write(strReturnStr &"<BR>") 
	Next 
	interpretaContratos=sContrato
End Function

Const MinNum = 0
Const MaxNum = 4294967295.99

Function Pesos(Number)
	DIM strPesos
	DIM CompletarDecimales
	If (Number >= MinNum) And (Number <= MaxNum) Then
		Pesos = conLetra(Fix(Number))
		If CSNG(Round((Number - Fix(Number)) * 100)) < 10 Then
			CompletarDecimales="0"
		Else
			CompletarDecimales=""
		End If
		IF Fix(Number)=1 THEN
			strPesos="PESO"
		ELSE 
			strPesos="PESOS"
		END IF
		Pesos = Pesos & " "& strPesos &" " & CompletarDecimales & CStr(Round((Number - Fix(Number)) * 100)) & "/100 M.N."
	Else
		Pesos = "Error, verifique la cantidad."
	End If
End Function

Function conLetra(N)
Dim Numbers, Tenths, decimales, Hundrens
Dim Result, primeraParte_letra, separador, resto_letra
IF session("IdUsuario")=1 THEN
	decimales=Round((N - Fix(N)) * 100)
	Numbers = Array("CERO", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
	Tenths = Array("CERO", "DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
	Hundrens = Array("CERO", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
	IF decimales>0 THEN
		primeraParte_letra=conLetra(Fix(N))
		IF MID(primeraParte_letra, LEN(primeraParte_letra)+1-2, 2)="UN" THEN
			primeraParte_letra=primeraParte_letra&"O"
		END IF
		resto_letra=TRIM(conLetra(decimales))
		IF decimales<10 THEN 
			resto_letra=" CERO " & resto_letra
		END IF
		IF MID(resto_letra, LEN(resto_letra)+1-2, 2)="UN" THEN
			resto_letra=resto_letra&"O"
		END IF
		separador=" PUNTO "
	ELSEIF N=0 THEN
		primeraParte_letra=""
		separador=""
		resto_letra=""
	ELSEIF N>=1 AND N<30 THEN
		primeraParte_letra=Numbers(N)
		separador=""
		resto_letra=""
	ELSEIF N>=30 AND N<100 THEN
		primeraParte_letra=Tenths(N \ 10)
		If N Mod 10 <> 0 Then
			Separador=" Y "
		Else
			Separador=" "
		End If
		resto_letra=conLetra(N Mod 10)
	ELSEIF N>=100 AND N<1000 THEN
		primeraParte_letra=""
		separador=""
		resto_letra=""
		IF Fix(N \ 100) = 1 THEN
			primeraParte_letra = "CIEN"
		ELSE
			primeraParte_letra = Hundrens(N \ 100)
		END IF
		separador=" "
		resto_letra=conLetra(N Mod 100)
	ELSEIF N>=1000 AND N<1000000 THEN
		primeraParte_letra=conLetra(N \ 1000)
		separador=" MIL "
		resto_letra=conLetra(N Mod 1000)
	ELSEIF N>=1000000 AND N<1000000000 THEN
		primeraParte_letra = conLetra(N \ 1000000)
		IF Fix(N \ 1000000)=1 THEN
			separador = " MILLON "
		ELSE
			separador = " MILLONES "
		END IF
		resto_letra = conLetra(N Mod 1000000)
	ELSEIF N>=1000000000 AND N<MaxNum+1 THEN
		primeraParte_letra = conLetra(N \ 1000000000)
		IF Fix(N \ 1000000000)=1 THEN
			separador = " BILLON "
		ELSE
			separador = " BILLONES "
		END IF
		resto_letra = conLetra(N Mod 1000000000)
	END IF
	Result = primeraParte_letra + separador + resto_letra
ELSE
	decimales=Round((N - Fix(N)) * 100)
	Numbers = Array("CERO", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
	Tenths = Array("CERO", "DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
	Hundrens = Array("CERO", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
	IF decimales>0 THEN
		primeraParte_letra=conLetra(Fix(N))
		IF MID(primeraParte_letra, LEN(primeraParte_letra)+1-2, 2)="UN" THEN
			primeraParte_letra=primeraParte_letra&"O"
		END IF
		resto_letra=TRIM(conLetra(decimales))
		IF decimales<10 THEN 
			resto_letra=" CERO " & resto_letra
		END IF
		IF MID(resto_letra, LEN(resto_letra)+1-2, 2)="UN" THEN
			resto_letra=resto_letra&"O"
		END IF
		Result = primeraParte_letra + " PUNTO " + resto_letra
	ELSEIF N=0 THEN
		Result = ""
	ELSEIF N>=1 AND N<30 THEN
		Result = Numbers(N)
	ELSEIF N>=30 AND N<100 THEN
		If N Mod 10 <> 0 Then
			Result = Tenths(N \ 10) + " Y " + conLetra(N Mod 10)
		Else
			Result = Tenths(N \ 10) + " " + conLetra(N Mod 10)
		End If
	ELSEIF N>=100 AND N<1000 THEN
		If N \ 100 = 1 Then
			If N = 100 Then
				Result = "CIEN" + " " + conLetra(N Mod 100)
			Else
				Result = Hundrens(N \ 100) + " " + conLetra(N Mod 100)
			End If
		Else
			Result = Hundrens(N \ 100) + " " + conLetra(N Mod 100)
		End If
	ELSEIF N>=1000 AND N<1000000 THEN
		Result = conLetra(N \ 1000) + " MIL " + conLetra(N Mod 1000)
	ELSEIF N>=1000000 AND N<1000000000 THEN
		IF Fix(N \ 1000000)=1 THEN
			Result = conLetra(N \ 1000000) + " MILLON " + conLetra(N Mod 1000000)
		ELSE
			Result = conLetra(N \ 1000000) + " MILLONES " + conLetra(N Mod 1000000)
		END IF
	ELSEIF N>=1000000000 AND N<MaxNum+1 THEN
		Result = conLetra(N \ 1000000000) + " BILLONES " + conLetra(N Mod 1000000000)
	END IF
END IF
	conLetra = Result
End Function

 Function fncLetra(byVal cant)
if (cant=0) then
	strng="CERO "
else
	decimales=round((cant-fix(cant))*100)
	cant=fix(cant)
	strng=""
	temp=cant/1000000
	temp1=fix(temp)
	cant=cant-temp1*1000000
	if (temp1>0) then
		if (temp1>1) then
			strng=strng&fncLetra(temp1)& "MILLONES "
		else
			strng=strng&fncLetra(temp1)& "MILLÓN "
		end if
	end if

'	//con esto checamos los miles
	temp=cant/1000
	temp1=fix(temp)
	cant=cant-temp1*1000
	if (temp1>0) then
		strng=strng&fncLetra(temp1)& "MIL "
	end if
'	//con esto checamos las centenas
	
	temp=cant/100
	temp1=fix(temp)
	temp2=(cant-temp1*100)/10
	cant=cant-temp1*100
	if (temp1=1) then
		if (temp2<>0) then
			strng=strng&"CIENTO "
		else
			strng=strng&"CIEN "
		end if
	elseif (temp1=2) then
		strng=strng&"DOSCIENTOS "
	elseif (temp1=3) then
		strng=strng&"TRESCIENTOS "
	elseif (temp1=4) then
		strng=strng&"CUATROCIENTOS "
	elseif (temp1=5) then
		strng=strng&"QUINIENTOS "
	elseif (temp1=6) then
		strng=strng&"SEISCIENTOS "
	elseif (temp1=7) then
		strng=strng&"SETECIENTOS "
	elseif (temp1=8) then
		strng=strng&"OCHOCIENTOS "
	elseif (temp1=9) then
		strng=strng&"NOVECIENTOS "
	end if
	
'	//con esto checamos las decenas
	temp=cant/10
	temp1=fix(temp)
	temp2=cant-temp1*10
	if (temp1<3) then
		if (temp1=1) then
			if (temp2<6) then
				if (temp2=0) then
					strng=strng&"DIEZ "
				elseif (temp2=1) then
					strng=strng&"ONCE "
				elseif (temp2=2) then
					strng=strng&"DOCE "
				elseif (temp2=3) then
					strng=strng&"TRECE "
				elseif (temp2=4) then
					strng=strng&"CATORCE "
				elseif (temp2=5) then
					strng=strng&"QUINCE "
				end if
				temp2=0
			else
				strng=strng&"DIECI"
			end if
		elseif (temp1=2) then
			if (temp2=0) then
				strng=strng&"VEINTE "
			else
				strng=strng&"VEINTI"
			end if
		end if
	else
		if (temp1=3) then
			strng=strng&"TREINTA "
		elseif (temp1=4) then
			strng=strng&"CUARENTA "
		elseif (temp1=5) then
			strng=strng&"CINCUENTA "
		elseif (temp1=6) then
			strng=strng&"SESENTA "
		elseif (temp1=7) then
			strng=strng&"SETENTA "
		elseif (temp1=8) then
			strng=strng&"OCHENTA "
		elseif (temp1=9) then
			strng=strng&"NOVENTA "
		end if
	
		if (temp2<>0) then
			strng = strng & "Y "
		end if
	end if
		
'	//con esto checamos los demás
	if (temp2=1) then
		strng = strng & "UN "
	elseif (temp2=2) then
		strng = strng & "DOS "
	elseif (temp2=3) then
		strng = strng & "TRES "
	elseif (temp2=4) then
		strng = strng & "CUATRO "
	elseif (temp2=5) then
		strng = strng & "CINCO "
	elseif (temp2=6) then
		strng = strng & "SEIS "
	elseif (temp2=7) then
		strng = strng & "SIETE "
	elseif (temp2=8) then
		strng = strng & "OCHO "
	elseif (temp2=9) then
		strng = strng & "NUEVE "
	end if

	IF decimales>0 THEN
		strng = strng & "PUNTO "&fncLetra(decimales)
	END IF

end if
fncLetra=strng
End Function

FUNCTION updateURLString(ByVal strQueryString, ByVal variable, ByVal valor)
DIM new_string, variable_buscar, pos_variable, left_string, right_string, strPartial, pos_partial
strQueryString=TRIM(URLDecode(strQueryString))
'RESPONSE.WRITE strQueryString & "<br><br>"
'RESPONSE.END
new_string=replace(strQueryString, "?", "&")
variable_buscar="&"&variable&"="
pos_variable=INSTR(UCASE(new_string), UCASE(variable_buscar))
IF pos_variable>0 THEN
	left_string=MID(strQueryString, 1, pos_variable-1+LEN(variable_buscar)) & valor
	strPartial=RIGHT(strQueryString, LEN(strQueryString)-pos_variable-LEN(variable_buscar)+1)
	pos_partial=INSTR(strPartial, "&")
	IF pos_partial>0 THEN 
		right_string=RIGHT(strPartial, LEN(strPartial)-pos_partial+1)
	END IF
ELSE
	left_string=strQueryString & variable_buscar & valor
	right_string=""
END IF
new_string=REPLACE(left_string, ".asp&", ".asp?")&right_string
'new_string=REPLACE(UCASE(strQueryString), UCASE(variable), UCASE(variable)&"="&valor)

updateURLString=new_string
END FUNCTION 


function IsNaN(byval n) 'http://www.livio.net/main/asp_functions.asp?id=IsNaN%20Function
    dim d
    'ON ERROR  RESUME NEXT
    if not isnumeric(n) then
        IsNan = true
        Exit Function
    end if
    d = cdbl(n)
    if err.number <> 0 then isNan = true else isNan = false
	'Err.Clear
	'ON ERROR  GOTO 0
end function

Function Concatenate(byval sString1, byval sString2)
	Concatenate=sString1&sString2
End Function

Sub Assign(ByRef oTarget, ByRef vValue)
'	response.write "<br>"&[&TypeName](vValue)&"<br>"
	ON ERROR	RESUME NEXT
	IF IsObject(vValue) THEN
		Set oTarget = vValue
	ELSE
		oTarget = vValue
	END IF
	[&Catch] TRUE, Me, "Assign", "Error para "&vValue&""
	ON ERROR	GOTO 0
End Sub

Function IsNullOrEmpty(ByVal sInput)
	IF IsObject(sInput) THEN
		IsNullOrEmpty = FALSE
	ELSE
		ON ERROR  RESUME NEXT
		IF UCASE([&TypeName](sInput))="BYTE()" THEN
			IsNullOrEmpty =IsNull(sInput)
		ELSE
			IsNullOrEmpty = (IsNull(sInput) OR IsEmpty(sInput) OR sInput="")
		END IF
		IF Err.Number<>0 THEN
			RESPONSE.WRITE "Error con tipo de datos "&[&TypeName](sInput)
			Err.Clear
			RESPONSE.END
		END IF
		
		ON ERROR  GOTO 0
	END IF
End Function

Function IsBlankOrEmpty(ByVal sInput)
	IsBlankOrEmpty=( IsEmpty(sInput) OR TRIM(sInput)="" )
End Function

Function NullIfEmptyOrNullText(sInput)
	NullIfEmptyOrNullText=Try( ((sInput="NULL") OR IsNullOrEmpty(sInput)), NULL, sInput)
End Function

Function EmptyIfNull(sInput)
	IF IsNull(sInput) THEN
		EmptyIfNull=Empty
	ELSE
		EmptyIfNull=sInput
	END IF
End Function

Function ZeroIfInvalid(sInput)
	IF NOT ISNUMERIC(sInput) THEN
		ZeroIfInvalid=0
	ELSE
		ZeroIfInvalid=sInput
	END IF
End Function

Function AppendIfNotEmpty(ByVal sString, ByVal sAppendString, ByVal sPosition)
IF NOT(IsNullOrEmpty(sPosition)) THEN sPosition="after"
IF NOT(IsNullOrEmpty(sString)) THEN 
	IF UCASE(sPosition)="BEFORE" THEN
		sString=sAppendString&sString
	ELSE
		sString=sString&sAppendString
	END IF
END IF
AppendIfNotEmpty=sString
End Function 

Function IsObjectReady(oInput)
	IF NOT IsObject(oInput) THEN 
		Err.Raise 1, "ASP 101", "Input variable is not an object"
		response.end
	END IF
	IsObjectReady=NOT(oInput IS NOTHING)
End Function

Sub [&CheckValidObjectType](oInput, sValidDataTypes)
	IF NOT CString([&TypeName](oInput)).ExistsIn(TRIM(sValidDataTypes)) THEN
		RESPONSE.WRITE "<strong>"&[&TypeName](oInput)&"</strong> is not a valid type, object only admits "&sValidDataTypes&" Type Objects"
		RESPONSE.END
	END IF
End Sub

Function TryPropertyRemove(ByRef oElement, sPropertyName, ByVal WarnError)
	DIM bReturnValue
	ON ERROR  RESUME NEXT
'	IF session("debug") THEN _
'	response.write "<br><strong>TryPropertySet: </strong>Intentando propiedad <strong class=""info"">"&sPropertyName&"</strong> para <strong class=""info"">"&[&TypeName](oElement)&"</strong> "
	oElement.Properties.RemoveProperty sPropertyName

	IF Err.Number<>0 THEN
		IF WarnError OR session("debug") THEN 
			RESPONSE.WRITE "<strong class=""warning"">Warning: </strong>No se puede <strong class=""info"">establecer</strong>, para el tipo de objeto <strong class=""info"">"&[&TypeName](oElement)&"</strong>, la propiedad <strong class=""info"">"&sPropertyName&"</strong> ("&Try(IsObject(vInput), "Object: "&[&TypeName](vInput), vInput)&")"
			IF [&TypeName](oElement)="DataField" THEN 
				RESPONSE.WRITE " ("&[&TypeName](oElement.Control)&")"
			END IF
			RESPONSE.WRITE "("&Err.Description&")<br>"
		END IF
'		response.end
		Err.Clear
		bReturnValue=FALSE
	ELSE
		bReturnValue=TRUE
	END IF
	ON ERROR  GOTO 0
	TryPropertyRemove=bReturnValue
End Function

Function TryPropertySet(ByRef oElement, sProperties, ByRef vInput, ByVal WarnError)
	DIM bReturnValue, aPropertyNames
	ON ERROR  RESUME NEXT
'	IF sProperties="Control.IsReadonly" THEN _
	IF session("debug") THEN _
	Debugger Me, "<br><strong>TryPropertySet: </strong>Intentando propiedad <strong class=""info"">"&sProperties&"</strong> para <strong class=""info"">"&[&TypeName](oElement)&" Valor ("&TypeName(vInput)&"): "&vInput&"</strong> "
	IF INSTR(sProperties, "/")<>0 THEN 
		aPropertyNames = Split(sProperties, "/")
	ELSE
		aPropertyNames = Array(sProperties)
	END IF
	
	DIM sPropertyName
	DIM iPropertyName:	iPropertyName=0
	FOR EACH sPropertyName IN aPropertyNames
		iPropertyName=iPropertyName+1
		IF Err.Number<>0 OR iPropertyName=1 THEN
			Err.Clear
'			IF IsObject(vInput) THEN
'				EXECUTE("SET oElement."&sPropertyName&"=vInput")
'			ELSE
'				EXECUTE("oElement."&sPropertyName&"=vInput")
'			END IF
			EXECUTE("SET oElement."&sPropertyName&"=vInput")
			IF Err.Number<>0 THEN
				Err.Clear
				EXECUTE("oElement."&sPropertyName&"=vInput")
			END IF
		END IF
	NEXT
	
	IF Err.Number<>0 THEN
		IF WarnError OR session("debug") THEN 
			RESPONSE.WRITE "<strong class=""warning"">Warning: </strong>No se puede <strong class=""info"">establecer</strong>, para el tipo de objeto <strong class=""info"">"&[&TypeName](oElement)&"</strong>, la propiedad <strong class=""info"">"&sProperties&"</strong> ("&Try(IsObject(vInput), "Object: "&[&TypeName](vInput), vInput)&")"
			IF [&TypeName](oElement)="DataField" THEN 
				RESPONSE.WRITE " ("&[&TypeName](oElement.Control)&")"
			END IF
			RESPONSE.WRITE "("&Err.Description&")<br>"
		END IF
'		response.end
		Err.Clear
		bReturnValue=FALSE
	ELSE
		bReturnValue=TRUE
	END IF
	ON ERROR  GOTO 0
	TryPropertySet=bReturnValue
End Function
Sub [&Stop](oSource)
	response.write "<strong style=""color:'red'"">Detenido en "&[&TypeName](oSource)&"</strong>"
	response.end
End Sub

Function TryPropertyGet(ByRef oElement, sPropertyName, ByRef vTarget, ByVal WarnError)
	DIM bReturnValue
	ON ERROR  RESUME NEXT
	IF session("debug") THEN _
	response.write "<br><strong>TryPropertyGet: </strong>Intentando propiedad <strong class=""info"">"&sPropertyName&"</strong> para <strong class=""info"">"&[&TypeName](oElement)&"</strong> "
	
	EXECUTE("IF IsObject(oElement."&sPropertyName&") THEN Set vTarget=oElement."&sPropertyName&" ELSE vTarget=oElement."&sPropertyName&" END IF")
	IF Err.Number<>0 THEN
		IF WarnError OR session("debug") THEN RESPONSE.WRITE "<strong class=""warning"">Warning: </strong>La propiedad "&sPropertyName&" no se puede <strong class=""info"">recuperar</strong> para el tipo de objeto "&[&TypeName](oElement)&"<br>"
'		response.end
		Err.Clear
		bReturnValue=FALSE
	ELSE
		bReturnValue=TRUE
	END IF
	ON ERROR  GOTO 0
	TryPropertyGet=bReturnValue
End Function

Function fillWith(ByVal sValue, ByVal sFill, ByVal iLength, ByVal sPosition)
DIM sNewString
IF IsNullOrEmpty(sPosition) THEN sPosition="right"
DIM iCurrentLength: iCurrentLength=LEN(TRIM(sValue))
DIM sFilled: sFilled=replicate(sFill, iLength-iCurrentLength)
IF sPosition="left" THEN sNewString=sFilled&sNewString
sNewString=sNewString&sValue
IF sPosition="right" THEN sNewString=sNewString&sFilled
fillWith=sNewString
End Function

Function replicate(ByVal sString, ByVal iTimes)
DIM sNewString: sNewString=""
DO WHILE(iTimes>0) 
	sNewString=sNewString&sString
	iTimes=iTimes-1
LOOP
replicate=sNewString
End Function

Function QuoteName(sString)
	QuoteName = "'"&sString&"'"
End Function

Function DoubleQuoteName(sString)
	QuoteName = """"&sString&""""
End Function

Function ReferenciaHorizontes(sLote)
ReferenciaHorizontes=CONCATENATE(CSTR(2),fillWith(Lote,"0",3,"left"))
End Function

Function Try(bCondition, sTrue, sFalse)
	IF bCondition THEN
		Try=sTrue
	ELSE
		Try=sFalse
	END IF
End Function

Function [&Coalesce](vFirstOption, vSecondOption)
	IF NOT IsNullOrEmpty(vFirstOption) THEN
		[&Coalesce]=vFirstOption
	ELSE
		[&Coalesce]=vSecondOption
	END IF
End Function

Function ToArray(ByVal oObject)
	DIM oArray: Set oArray = new ArrayList
	DIM oElement
	FOR EACH oElement IN oObject
		oArray.Add oElement
	NEXT
	ToArray = oArray.ToArray()
End Function

Function TextToNull(ByVal vValue)
	TextToNull=Try(vValue="NULL", NULL, vValue)
End Function

Function NullToText(ByVal vValue)
	IF IsNull(vValue) THEN 
		NullToText="NULL"
	ELSE
		NullToText=vValue
	END IF
End Function

Function ClassTracker(oClass)
	ClassTracker="<strong class=""warning"">"& [&TypeName](oClass)&" ==></strong>"
End Function

Sub Debugger(oClass, sText)
	DIM sType:
	IF [&TypeName](sText)="Byte()" THEN
		response.write ClassTracker(oClass)&" Dato de tipo :"& [&TypeName](sText)&"<br>"
	ELSE
		response.write ClassTracker(oClass)&" "&sText&"<br>"
	END IF
End Sub

Function [&TypeName](oInput)
ON ERROR	RESUME NEXT
	DIM sTypeName
	sTypeName=TypeName(oInput)
	IF Err.Number<>0 THEN
		Err.Clear
		Debugger Me, "<strong>[&TypeName]</strong>"&oInput.Type
		sTypeName=oInput.Type
	END IF
[&TypeName]=sTypeName
ON ERROR	GOTO 0

End Function

Dim dBooleanDictionary
Set dBooleanDictionary = server.CreateObject("Scripting.Dictionary")
dBooleanDictionary("NO")="0"
dBooleanDictionary("YES")="1"
dBooleanDictionary(LCASE(CSTR(CBOOL(0))))="0" 'Traslates "False" OR "Falso" according to the language and assigns its value to 0
dBooleanDictionary(LCASE(CSTR(CBOOL(1))))="1"	'Traslates "True" OR "Verdadero" according to the language and assigns its value to 1
dBooleanDictionary("0")="0"
dBooleanDictionary("1")="1"
dBooleanDictionary("javascript_0")="false"
dBooleanDictionary("javascript_1")="true"

Function CBoolean(ByVal vValue)
	DIM oExtendedBoolean: Set oExtendedBoolean = new ExtendedBoolean
	oExtendedBoolean.Value=vValue
	Set CBoolean = oExtendedBoolean
End Function
Class ExtendedBoolean
	Private vValue
	Private dMainDictionary
	
	Private Sub Class_Initialize()
	End Sub
	Private Sub Class_Terminate()
	End Sub
	
	Public Default Property Get Value()
		Value = CBOOL(vValue)
	End Property
	Public Property Let Value(input)
		vValue = dBooleanDictionary(LCASE(CSTR(input)))
	End Property
	
	Public Function IsTrue()
		IsTrue=(Me.Value=TRUE)
	End Function
	
	Public Function IsFalse()
		IsFalse=(Me.Value=FALSE)
	End Function
	
	Public Function Text()
		Text=dBooleanDictionary("javascript_"&vValue)
	End Function
End Class

Public Function ChangeUnrecognizedFieldType(oField)
	DIM oResult
	SELECT CASE oField.Type
	CASE 20	'20: BigInt -- no soporta [&TypeName]
		oField.Type=3
	CASE ELSE
	END SELECT
	Set ChangeUnrecognizedFieldType=oField
End Function	

Function CString(ByVal input)
	DIM sText
	DIM oExtendedString: Set oExtendedString = new ExtendedString
	ON ERROR  RESUME NEXT
	IF UCASE([&TypeName](input))="FIELD" THEN 
		IF input.type=20 THEN '20: BigInt -- no soporta [&TypeName]
			input=CINT(input.value)
		ELSE
			input=input.value
		END IF
	END IF
	[&Catch] TRUE, Me, "CString", ""
	ON ERROR  GOTO 0

	SELECT CASE UCASE([&TypeName](input))
	CASE "NULL"
		sText=NULL
	CASE ELSE
		sText=CSTR(input) 'HTMLDecode(CSTR(input))
	END SELECT
	oExtendedString.Text=sText
	Set CString = oExtendedString
End Function

Class ExtendedString
	Private sText
	Public Default Property Get Text()
		Text = sText
	End Property
	Public Property Let Text(input)
		sText = input
	End Property

	Public Function RemoveEntities()
		Set RemoveEntities=Me.Replace("["&chr(13)&""&chr(9)&""&chr(10)&""&vbcr&""&vbcrlf&""&vbtab&"]", " ")
	End Function
	
	Public Function DoubleQuote()
		Me.Text=Me.Replace("""", """""").Append("""").Prepend("""")
		Set DoubleQuote = Me
	End Function 
	
	Public Function ExistsIn(sString)
		ExistsIn=(INSTR(" "&sString&",", " "&sText&",")>0)
	End Function

	Public Function IsLike(sPath)
		IsLike=testMatch(sText, sPath)
	End Function

	Public Function Remove(sPatrn)
		Set Remove=Me.Replace(sPatrn, "")
	End Function

	Public Function Replace(sPatrn, sReplacementText)
		Me.Text=replaceMatch(sText, sPatrn, sReplacementText)
		Set Replace = Me
	End Function

	Public Function ReplaceAndEvaluate(sPatrn, sReplacementText)
		Me.Text=replaceEvaluatingMatch(sText, sPatrn, sReplacementText)
		Set ReplaceAndEvaluate = Me
	End Function

	Public Function Escape()
		Me.Text=escapeString(sText)
		Set Escape = Me
	End Function

	Public Function EscapeChars(sChars)
		Me.Text=escapeCharacters(sText, sChars)
		Set EscapeChars = Me
	End Function

	Public Function Evaluate()
		Me.Text=fncEvaluate(sText)
		Set Evaluate = Me
	End Function

	Public Function Append(sAppend)
		Me.Text=sText&sAppend
		Set Append = Me
	End Function

	Public Function Prepend(sPrepend)
		Me.Text=sPrepend&sText
		Set Prepend = Me
	End Function

	Public Function IsNull()
		IsNull=CheckNull(sText)
	End Function
End Class

Public Function [&CString](ByVal input)
	DIM sText
	DIM oExtendedString: Set oExtendedString = new XString
	ON ERROR  RESUME NEXT
	IF UCASE([&TypeName](input))="FIELD" THEN 
		IF input.type=20 THEN '20: BigInt -- no soporta [&TypeName]
			input=CINT(input.value)
		ELSE
			input=input.value
		END IF
	END IF
	[&Catch] TRUE, Me, "[&CString]", ""
	ON ERROR  GOTO 0

	SELECT CASE UCASE([&TypeName](input))
	CASE "NULL"
		sText=NULL
	CASE ELSE
		sText=CSTR(input) 'HTMLDecode(CSTR(input))
	END SELECT
	oExtendedString.Text=sText
	Set [&CString] = oExtendedString
End Function	
Class XString
	Private oStringBuilder
	Private Sub Class_Initialize()
'		Set oStringBuilder = CreateObject("System.IO.StringWriter")		'Necesitaría .GetStringBuilder()
		Set oStringBuilder = CreateObject("System.Text.StringBuilder")	'Es más eficiente!!
	End Sub
	Private Sub Class_Terminate()
		Set oStringBuilder = nothing
	End Sub
	
	Public Default Property Get Text()
		Text = oStringBuilder.ToString()
	End Property
	Public Property Let Text(input)
		oStringBuilder.Length=0
		oStringBuilder.Append_3 CStr(input)
	End Property

	Public Function RemoveEntities()
		Set RemoveEntities=Me.Replace("["&chr(13)&""&chr(9)&""&chr(10)&""&vbcr&""&vbcrlf&""&vbtab&"]", " ")
	End Function
	
	Public Function DoubleQuote()
		Me.Text=Me.Replace("""", """""").Append("""").Prepend("""")
		Set DoubleQuote = Me
	End Function 
	
	Public Function ExistsIn(sString)
		ExistsIn=(INSTR(" "&sString&",", " "&Me.Text&",")>0)
	End Function

	Public Function IsLike(sPath)
		IsLike=testMatch(Me.Text, sPath)
	End Function

	Public Function Remove(sPatrn)
		Set Remove=Me.Replace(sPatrn, "")
	End Function

	Public Function RemoveString(sString)
		oStringBuilder.Replace sString, ""
		Set RemoveString = Me
	End Function

	Public Function Extract(sPatrn)
		Set Extract = [$ArrayList](getMatch(Me.Text, sPatrn))
	End Function
	
	Public Function Replace(sPatrn, sReplacementText)
		Me.Text=replaceMatch(Me.Text, sPatrn, sReplacementText)
		Set Replace = Me
	End Function

	Public Function Escape()
		Me.Text=escapeString(sText)
		Set Escape = Me
	End Function

	Public Function EscapeChars(sChars)
		Me.Text=escapeCharacters(sText, sChars)
		Set EscapeChars = Me
	End Function

	Public Function Evaluate()
		Me.Text=fncEvaluate(Me.Text)
		Set Evaluate = Me
	End Function

	Public Function Append(sAppend)
		oStringBuilder.Append_3 CStr(sAppend)
		Set Append = Me
	End Function

	Public Function Prepend(sPrepend)
'		oStringBuilder.AppendFormat_4 "Probando {0} en {1}", Array(1, "loneliest")
		oStringBuilder.Insert_2 0, CStr(sPrepend)
		Set Prepend = Me
	End Function

	Public Function IsNull()
		IsNull=CheckNull(sText)
	End Function
End Class

Function CheckNull(ByVal vValue)
	CheckNull=ISNULL(vValue)
End Function

Function escapeCharacters(sOriginal, sPatrn)
	DIM sReplacementText: sReplacementText="\$&"
	escapeCharacters=replaceMatch(sOriginal, sPatrn, sReplacementText)
End Function 

Function escapeString(sOriginal)
	Dim sPatrn: sPatrn="\[\\\^\$\.\|\?\*\+\(\)\{\}"
	escapeString=escapeCharacters(sOriginal, sPatrn)
End Function

Sub StoreSessionVariable(ByVal VariableName, ByVal vValue)
	IF IsObject(vValue) THEN
		Set Session(VariableName) = vValue
	ELSE
		Session(VariableName) = vValue
	END IF
End Sub 
 
Sub [&Catch](bStop, oSourceClass, sFunctionName, sMessage)
	IF Err.Number<>0 THEN
		response.write "<br>Ocurrió el siguiente error:<strong>"&Err.Description&"</strong>. En función:<strong>"&vbcrlf&TypeName(oSourceClass)&"."&sFunctionName&"</strong>.<br>"&sMessage&"<br><br>"
		IF bStop THEN response.end
		Err.Clear
	END IF
End Sub

Class Parameter
	Private sName, vValue
	Public Property Get Name()
		Name = sName
	End Property
	Public Property Let Name(input)
		sName = input
	End Property

	Public Property Get Value()
		Value = vValue
	End Property
	Public Property Let Value(input)
		vValue = input
	End Property
End Class

Class Parameters
	Private oDictionary
	Private Sub Class_Initialize()
		Set oDictionary = new Dictionary
	End Sub
	Private Sub Class_Terminate()
		Set oDictionary = nothing
	End Sub
	
	Public Property Get All()
		Set All = oDictionary
	End Property

	Public Default Property Get Parameters()
		Set Parameters = oDictionary
	End Property
	Public Property Let Parameters(input)
		DIM aControlParameters: Set aControlParameters=getParameters(input)
		DIM oParam
		FOR EACH oParam IN aControlParameters
			DIM oParameter: Set oParameter = new Parameter
			oParameter.Name=oParam.SubMatches(0)
			DIM oValue: Set oValue=[&CString](oParam.SubMatches(1))
			IF oValue.IsLike("^'.*'$") THEN
				oParameter.Value=oValue.Replace("\\'", """""").Replace("'", """").Text
			ELSE
				oParameter.Value=oValue.Text
			END IF
			oDictionary.Add oParameter.Name, oParameter
		NEXT
	End Property
End Class

Function [&New](ByVal vValue, ByVal sParameters)
DIM oCreated
Set oCreated=[&CreateObject](vValue)
IF NOT IsNullOrEmpty(sParameters) THEN
	DIM oParameters: Set oParameters = new Parameters
	oParameters.Parameters=sParameters
	DIM oParameter
	FOR EACH oParameter IN oParameters.All.Items
		'Set oParameter=oParameters.Parameters.Item(i)
	'	response.write [&TypeName](oParameter)
		EXECUTE("oCreated."&oParameter.Name&"="&oParameter.Value)
	'	RESPONSE.WRITE oParameter.Name&": "&oParameter.Value
	'	TryPropertySet oCreated, oParameter.Name, Evaluate(oParameter.Value), TRUE
	NEXT
	Set oParameters = NOTHING
END IF
Set [&New] = oCreated
End Function

Function [&CreateObject](ByVal sObject)
DIM oCreated
EXECUTE("Set oCreated = new "&sObject)
Set [&CreateObject]=oCreated
End Function 

Sub [&BR]()
	response.write "<br>"
End Sub

Sub [&echo](ByVal sText)
	response.write sText
End Sub

Class clsXML
  'strFile must be full path to document, ie C:\XML\XMLFile.XML
  'objDoc is the XML Object
  Private strFile, objDoc

  '*********************************************************************
  ' Initialization/Termination
  '*********************************************************************

  'Initialize Class Members
  Private Sub Class_Initialize()
    strFile = ""
  End Sub

  'Terminate and unload all created objects
  Private Sub Class_Terminate()
    Set objDoc = Nothing
  End Sub

  '*********************************************************************
  ' Properties
  '*********************************************************************

  'Set XML File and objDoc
  Public Property Let File(str)
    Set objDoc = Server.CreateObject("Microsoft.XMLDOM")
    objDoc.async = False
    strFile = str
    objDoc.Load strFile
  End Property

  'Get XML File
  Public Property Get File()
    File = strFile
  End Property

  '*********************************************************************
  ' Functions
  '*********************************************************************

  'Create Blank XML File, set current obj File to newly created file
  Public Function createFile(strPath, strRoot)
    Dim objFSO, objTextFile
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.CreateTextFile(strPath, True)
    objTextFile.WriteLine("<?xml version=""1.0""?>")
    objTextFile.WriteLine("<" & strRoot & "/>")
    objTextFile.Close
    Me.File = strPath
    Set objTextFile = Nothing
    Set objFSO = Nothing
  End Function

  'Get XML Field(s) based on XPath input from root node
  Public Function getField(strXPath)
    Dim objNodeList, arrResponse(), i
    Set objNodeList = objDoc.documentElement.selectNodes(strXPath)
    ReDim arrResponse(objNodeList.length)
    For i = 0 To objNodeList.length - 1
      arrResponse(i) = objNodeList.item(i).Text
    Next
    getField = arrResponse
  End Function

  'Update existing node(s) based on XPath specs
  Public Function updateField(strXPath, strData)
    Dim objField
    For Each objField In objDoc.documentElement.selectNodes(strXPath)
      objField.Text = strData
    Next
    objDoc.Save strFile
    Set objField = Nothing
    updateField = True
  End Function

  'Create node directly under root
  Public Function createRootChild(strNode)
    Dim objChild
    Set objChild = objDoc.createNode(1, strNode, "")
    objDoc.documentElement.appendChild(objChild)
    objDoc.Save strFile
    Set objChild = Nothing
  End Function

  'Create a child node under root node with attributes
  Public Function createRootNodeWAttr(strNode, attr, val)
    Dim objChild, objAttr
    Set objChild = objDoc.createNode(1, strNode, "")
    If IsArray(attr) And IsArray(val) Then
      If UBound(attr)-LBound(attr) <> UBound(val)-LBound(val) Then
        Exit Function
      Else
        Dim i
        For i = LBound(attr) To UBound(attr)
          Set objAttr = objDoc.createAttribute(attr(i))
          objChild.setAttribute attr(i), val(i)
        Next
      End If
    Else
      Set objAttr = objDoc.createAttribute(attr)
      objChild.setAttribute attr, val
    End If
    objDoc.documentElement.appendChild(objChild)
    objDoc.Save strFile
    Set objChild = Nothing
  End Function

  'Create a child node under the specified XPath Node
  Public Function createChildNode(strXPath, strNode)
    Dim objParent, objChild
    For Each objParent In objDoc.documentElement.selectNodes(strXPath)
      Set objChild = objDoc.createNode(1, strNode, "")
      objParent.appendChild(objChild)
    Next
    objDoc.Save strFile
    Set objParent = Nothing
    Set objChild = Nothing
  End Function

  'Create a child node(s) under the specified XPath Node with attributes
  Public Function createChildNodeWAttr(strXPath, strNode, attr, val)
    Dim objParent, objChild, objAttr
    For Each objParent In objDoc.documentElement.selectNodes(strXPath)
      Set objChild = objDoc.createNode(1, strNode, "")
      If IsArray(attr) And IsArray(val) Then
        If UBound(attr)-LBound(attr) <> UBound(val)-LBound(val) Then
          Exit Function
        Else
          Dim i
          For i = LBound(attr) To UBound(attr)
            Set objAttr = objDoc.createAttribute(attr(i))
            objChild.SetAttribute attr(i), val(i)
          Next
        End If
      Else
        Set objAttr = objDoc.createAttribute(attr)
        objChild.setAttribute attr, val
      End If
      objParent.appendChild(objChild)
    Next
    objDoc.Save strFile
    Set objParent = Nothing
    Set objChild = Nothing
  End Function

  'Delete the node specified by the XPath
  Public Function deleteNode(strXPath)
    Dim objOld
    For Each objOld In objDoc.documentElement.selectNodes(strXPath)
      objDoc.documentElement.removeChild objOld
    Next
    objDoc.Save strFile
    Set objOld = Nothing
  End Function
End Class

FUNCTION firstDayOfMonth(dDate)
	firstDayOfMonth=DATEADD("d", -DAY(dDate)+1, dDate)
END FUNCTION

FUNCTION lastDayOfMonth(dDate)
	lastDayOfMonth=DATEADD("d", -1, DATEADD("m", 1, firstDayOfMonth(dDate)))
END FUNCTION

Function JSONToXML(jsonString)
    Dim scriptControl, jsonObject, intermediateXML, finalXML, xsltRawToXML, xsltPrettifyJSON
    ' Note: VBScript regex is limited compared to SQL Server's, so use the Microsoft VBScript Regular Expressions object
    intermediateXML = TRIM(jsonString)

    Dim regex, matches, match
    ' Initialize regex object
    Set regex = New RegExp
    regex.Global = True
	regex.Pattern = "[ \r\n]+$"
    intermediateXML = regex.Replace(intermediateXML, "")

    ' Step 2: Replace JSON special characters to form intermediate XML-like format
    intermediateXML = Replace(intermediateXML, Chr(9), "<t/>") ' Replace tab characters
    intermediateXML = Replace(intermediateXML, Chr(10) & Chr(13), "<r/>") ' Replace newline characters
    intermediateXML = Replace(intermediateXML, Chr(13), "<r/>") ' Replace carriage returns
    intermediateXML = Replace(intermediateXML, ",", "<c/>") ' Replace commas
    intermediateXML = Replace(intermediateXML, "&", "&amp;") ' Replace ampersands

    ' Step 2: Perform regex replacements

    ' Replace \(.)
    regex.Pattern = "\\(.)"
    intermediateXML = regex.Replace(intermediateXML, "<e>$1</e>")

    intermediateXML = Replace(intermediateXML, "[", "<l>")
    intermediateXML = Replace(intermediateXML, "]", "</l>")
    intermediateXML = Replace(intermediateXML, "{", "<o>")
    intermediateXML = Replace(intermediateXML, "}", "</o>")
    intermediateXML = Replace(intermediateXML, " ", "<s/>")

    ' Replace "([^"]+?)":\s*
    regex.Pattern = """([^""]+?)"":\s*"
    intermediateXML = regex.Replace(intermediateXML, "<a>$1</a>")

    ' Replace <l>([^<]+)</l>
    regex.Pattern = "<l>([^<]+)</l>"
    intermediateXML = regex.Replace(intermediateXML, "<l>$1</l>")

    ' Step 3: Apply the raw-to-XML XSLT transformation
    Set xsltRawToXML = Server.CreateObject("MSXML2.FreeThreadedDOMDocument")
    xsltRawToXML.async = False
    xsltRawToXML.load(server.MapPath(".")&"\json_to_raw_xml.xslt")
    Set finalXML = Server.CreateObject("MSXML2.DOMDocument")
    finalXML.async = False
    finalXML.loadXML(intermediateXML)
    finalXML.loadXML(finalXML.transformNode(xsltRawToXML))

    ' Step 4: Apply the prettify-JSON XSLT transformation
    Set xsltPrettifyJSON = Server.CreateObject("MSXML2.FreeThreadedDOMDocument")
    xsltPrettifyJSON.async = False
    xsltPrettifyJSON.load(server.MapPath(".")&"\json_to_xml.xslt")
    finalXML.loadXML(finalXML.transformNode(xsltPrettifyJSON))

    ' Normalize namespaces and return final XML
    finalXML.setProperty "SelectionNamespaces", "xmlns:xsl='http://www.w3.org/1999/XSL/Transform' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xson='http://panax.io/xson'"
    'finalXML.normalizeNamespaces
    Set JSONToXML = finalXML

    ' Clean up
    Set scriptControl = Nothing
    Set jsonObject = Nothing
    Set xsltRawToXML = Nothing
    Set xsltPrettifyJSON = Nothing
    Set finalXML = Nothing
End Function

function apiCall(strUrl)
    Set xmlHttp = Server.Createobject("MSXML2.ServerXMLHTTP")
    xmlHttp.Open "GET", strUrl, False
    xmlHttp.setRequestHeader "User-Agent", "asp httprequest"
    xmlHttp.setRequestHeader "content-type", "application/x-www-form-urlencoded"
    xmlHttp.Send
    'response.write xmlHttp.responseText
    'xmlHttp.abort()
	apiCall = xmlhttp.ResponseText
    set xmlHttp = Nothing   
end function 

function asyncCall(strUrl)
    Set xmlHttp = Server.Createobject("MSXML2.ServerXMLHTTP")
    xmlHttp.Open "GET", strUrl, False
    xmlHttp.setRequestHeader "User-Agent", "asp httprequest"
    xmlHttp.setRequestHeader "content-type", "application/x-www-form-urlencoded"
    xmlHttp.Send
    'response.write xmlHttp.responseText
    'xmlHttp.abort()
    set xmlHttp = Nothing   
end function 

function Sleep(seconds)
    set oShell = CreateObject("Wscript.Shell")
    cmd = "%COMSPEC% /c timeout " & seconds & " /nobreak"
    oShell.Run cmd,0,1
End function

function curPageURL()
 dim protocol, port
 protocol = LCase(Request.ServerVariables("SERVER_PROTOCOL"))
 protocol=Left(protocol, instrRev(protocol, "/")-1)
 if Request.ServerVariables("HTTPS") = "on" then
   protocol=protocol&"s"
 end if  

 if Request.ServerVariables("SERVER_PORT") = "80" then
   port = ""
 else
   port = ":" & Request.ServerVariables("SERVER_PORT")
 end if  

 curPageURL = protocol & "://" & Request.ServerVariables("SERVER_NAME") &_
              port & Request.ServerVariables("SCRIPT_NAME")
end function

Dim SqlRegEx: Set SqlRegEx = New RegExp
With SqlRegEx
    .Pattern = "(\[[^\[]*\])+"
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
End With

Dim RegEx_JS_Escape: Set RegEx_JS_Escape = New RegExp
With RegEx_JS_Escape
    .Pattern = """|\\"
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
End With

Function login()
	Dim oCn: Set oCn = Server.CreateObject("ADODB.Connection")
	oCn.ConnectionTimeout = 5
	oCn.CommandTimeout = 180

    Response.CharSet = "ISO-8859-1"
    DIM oConfiguration:	set oConfiguration = Server.CreateObject("MSXML2.DOMDocument"): 
    oConfiguration.Async = false: 
    oConfiguration.setProperty "SelectionLanguage", "XPath"
    oConfiguration.Load(Server.MapPath("../system.config"))
	IF oConfiguration.documentElement IS NOTHING THEN
		oConfiguration.Load(Server.MapPath("../../.config/system.config"))
		IF oConfiguration.documentElement IS NOTHING THEN
			oConfiguration.Load(Server.MapPath("../../config/system.config"))
			IF oConfiguration.documentElement IS NOTHING THEN
				oConfiguration.Load(Server.MapPath("../../../.config/system.config"))
			END IF
		END IF
	END IF

    DIM sConnectionId
    IF  request.form("database_id")<>"" THEN
	    sConnectionId=request.form("database_id")
    ELSEIF Application("database_id")<>"" THEN
	    sConnectionId=Application("database_id")
    ELSEIF  request.form("Connection_id")<>"" THEN
	    sConnectionId=request.form("Connection_id")
    ELSEIF Application("Connection_id")<>"" THEN
	    sConnectionId=Application("Connection_id")
	ELSE
		sConnectionId=request.serverVariables("HTTP_HOST")
    END IF

    IF oConfiguration.documentElement IS NOTHING THEN
        Response.ContentType = "application/json"
        Response.CharSet = "ISO-8859-1"
        Response.Status = "412 Precondition Failed" %>
	    {
	    "success": false,
	    "message": "No se encontró el archivo de configuración system.config"
	    }
    <% 	response.end
    END IF

    DIM sConnectionString
    IF sConnectionId<>"" THEN
	    sConnectionString="@Id='"&sConnectionId&"' or Alias/text()='"&sConnectionId&"'"
    ELSE
	    sConnectionString="1=0"
    END IF

    DIM oDatabase: 
    SET oDatabase = oConfiguration.documentElement.selectSingleNode("/configuration/Databases/*["&sConnectionString&"]")
    IF oDatabase IS NOTHING THEN
        SET oDatabase = oConfiguration.documentElement.selectSingleNode("/configuration/Databases/*[@Id=../@Default or string(../@Default)='' and (@Id='default' or @Id='main')]")
    END IF

    IF oDatabase IS NOTHING THEN
        Response.ContentType = "application/json"
        Response.CharSet = "ISO-8859-1"
        Response.Status = "401 Unauthorized" %>
	    {
	    "success": false,
	    "message": "No se encontró definida la conexión <%= REPLACE(sConnectionId,"\","\\") %> en el archivo de configuración system.config"
	    }
    <% 	response.end
    END IF
    SESSION("connection_id") = oDatabase.getAttribute("Id")
    SESSION("database_id") = oDatabase.getAttribute("Id")

    DIM sDatabaseName, sDatabaseDriver, sDatabaseEngine, sDatabaseServer, sDatabaseUser, sDatabasePassword
    sDatabaseName  		= oDatabase.getAttribute("Name")
    sDatabaseEngine 	= oDatabase.getAttribute("Engine")
    sDatabaseServer		= oDatabase.getAttribute("Server")
    sDatabaseUser     	= oDatabase.getAttribute("User")
    sDatabasePassword 	= oDatabase.getAttribute("Password")
    sDefaultUser     	= oDatabase.getAttribute("DefaultUser")
    IF ISNULL(sDefaultUser) THEN
        sDefaultUser     	= ""
    END IF
	DIM authorization: authorization = Request.ServerVariables("HTTP_AUTHORIZATION")
	DIM decrypted_password
	If authorization<>"" Then
		DIM DecodedAuthorization: DecodedAuthorization = Base64Decode(MID(authorization,7))
		sUserLogin = Split(DecodedAuthorization, ":")(0)
		sUserName = sUserLogin
		decrypted_password = Split(DecodedAuthorization, ":")(1)
		If LEN(decrypted_password) = 32 OR LEN(decrypted_password) >= 1000 OR LEN(decrypted_password) = 0 then
			sPassword = decrypted_password
		Else
			sPassword = Hash("md5",decrypted_password)
		End If
	Else
		sUserLogin = LCASE(URLDecode(request.form("UserName")))
		sUserName = sUserLogin
		sPassword = URLDecode(request.form("Password"))
	End if
	session("user_login") = sUserName

	DIM oUser
	SET oUser=oDatabase.selectSingleNode("(./User[@Name='"&sUserName&"' or not(../User[@Name='"&sUserName&"']) and (@Name='*' or starts-with(@Name,'*@') and contains('"&sUserName&"',substring(@Name,3)))])[last()]")
	Set rsResult = Server.CreateObject("ADODB.RecordSet")
    SESSION("secret_engine") = sDatabaseEngine
	IF oUser IS NOTHING THEN
		Response.ContentType = "application/json"
		Response.CharSet = "ISO-8859-1"
		Response.Status = "401 Unauthorized" %>
		{
		"success": false,
		"message": "Usuario no autorizado"
		}
<% 	    response.end
	END IF
    IF LCASE(SESSION("secret_engine")) = "google" THEN
		session("secret_token") = sPassword
		checkConnection(oCn)
		Session("AccessGranted") = TRUE
		session("status") = "authorized"

		' Define the structure of the recordset (fields)
		rsResult.Fields.Append "user_id", adInteger
		rsResult.Fields.Append "user_name", adVarChar, 255

		' Open the recordset for editing
		rsResult.Open

		' Add records to the recordset
		rsResult.AddNew
		rsResult("user_id").Value = 99999
		rsResult("user_name").Value = sUserName
		rsResult.Update
		'response.write "strSQL: "&strSQL: response.end

		Set Login = rsResult
		Exit Function
	ELSEIF ISNULL(sDatabaseUser) THEN
        IF sUserName="" AND sPassword="" AND sDefaultUser<>"" THEN
            sUserName = sDefaultUser
        END IF

		sDatabaseUser = oUser.getAttribute("InstanceUser")
		IF ISNULL(sDatabaseUser) THEN
			sDatabaseUser = sUserName
		END IF
		IF sPassword="" AND NOT ISNULL(oUser.getAttribute("Password")) THEN
			sPassword = oUser.getAttribute("Password")
		END IF
		IF NOT ISNULL(oUser.getAttribute("InstancePassword")) THEN
			sDatabasePassword 	= oUser.getAttribute("InstancePassword")
		END IF
	ELSE
        IF sUserName<>"webmaster" THEN
            sDatabaseUser = sUserName
            'sDatabasePassword = "40A965D05136639974C40FAF6CFDF21D"
            'IF sUserName="guest" THEN
            '    sPassword = "40A965D05136639974C40FAF6CFDF21D"
            'END IF
        END IF
		sUserLogin = LCASE(URLDecode(request.form("UserName")))
		sUserName = sUserLogin
		sPassword = URLDecode(request.form("Password"))
    END IF

	IF ISNULL(sDatabasePassword) THEN
		sDatabasePassword = decrypted_password
	END IF

    SESSION("secret_database_user") = sDatabaseUser
    SESSION("secret_database_password") = sDatabasePassword
    SESSION("secret_server_id") = oDatabase.getAttribute("Server")
    SESSION("secret_database_name") = sDatabaseName

    DIM currentLocation: currentLocation = curPageURL()
    
	ON ERROR RESUME NEXT
    checkConnection(oCn)
	IF oCn.state=0 THEN
		Set Login = nothing
	ELSE
		session("user_login") = sUserName
		strSQL="EXEC [#Security].Authenticate '" & REPLACE(RTRIM(sUserName),"'", "''") & "', '"& REPLACE(RTRIM(sPassword),"'", "''") & "'"
		'response.write "strSQL: "&strSQL: response.end
		rsResult.CursorLocation 	= 3
		rsResult.CursorType 		= 3
		set rsResult = oCn.Execute(strSQL)
		IF Err.Number<>0 THEN 
			Session("AccessGranted") = FALSE
			session("status") = "unauthorized"
			Response.ContentType = "application/json"
			Response.CharSet = "ISO-8859-1"
			IF Err.Number=-2147217911 THEN
				Response.Status = "401 Unauthorized"
			ELSE 
				Response.Status = "409 Conflict"
			END IF
		ELSE
			Session("AccessGranted") = TRUE
			session("status") = "authorized"
		END IF
		checkConnection(oCn)
		'	alert('<%= REPLACE(strSQL, "'", "\'") %%')
		'<%	'response.end
		'Response.CodePage = 65001
		'Response.CharSet = "UTF-8"
		Set Login = rsResult
	END IF
End Function
    %>