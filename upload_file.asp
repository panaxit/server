<!--#include file="vbscript.asp"-->
<% SERVER.SCRIPTTIMEOUT = 4800 %>
<%
ON ERROR RESUME NEXT
'Stores only files with size less than MaxFileSize

'Using Huge-ASP file upload
'Dim Form: Set Form = Server.CreateObject("ScriptUtils.ASPForm")
'Using Pure-ASP file upload
Dim Form: Set Form = New ASPForm %>
<!--#INCLUDE FILE="upload.motobit.asp"-->
<% 
Server.ScriptTimeout = 2000
Form.SizeLimit = &HF000000

'{b}Set the upload ID for this form.
'Progress bar window will receive the same ID.
Dim uploadId: uploadId = Request.QueryString("UploadID")
IF uploadId="" THEN
	uploadId=Request.ServerVariables("HTTP_X_UPLOAD_ID")
END IF
if len(UploadID)>0 then
	Form.UploadID = uploadId'{/b}
    else
    Randomize
	Form.UploadID = clng(rnd * &H7FFFFFFF)
end if
'was the Form successfully received?
Const fsCompleted  = 0
If Form.State = fsCompleted Then 'Completed
  'was the Form successfully received?
  If Form.State = 0 then %><% 
	Dim parent_folder
	Dim fileName, saveAs, extension
	Dim relativeTargetPath, absolute_path
	Dim fso:	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	parent_folder=Request.QueryString("parentFolder")'TRIM(Form("parentFolder").Value)
	IF parent_folder="" THEN
		parent_folder=Request.ServerVariables("HTTP_X_PARENT_FOLDER")
	END IF
  IF parent_folder="" THEN 
    parent_folder="FilesRepository"
  END IF
	IF NOT fso.FolderExists(parent_folder_path) THEN
    Err.Raise vbObjectError + 1000, "uploadFile", "Parent folder is not authorized or does not exist"
	END IF
%>
		    {
		    "files": [{}
<%  DIM File 
		Response.ContentType = "application/json"
    FOR EACH File IN Form.Files.Items 
      saveAs=Request.QueryString("saveAs")'TRIM(Form("saveAs").Value)
		  IF saveAs = "" THEN
			  saveAs=Request.ServerVariables("HTTP_X_SAVE_AS")
		  END IF

      extension = fso.GetExtensionName(Form.Files.Item(File.Name).FileName)

      file_path = saveAs
      IF TRIM(file_path) = "" THEN
          file_path = File.FilePath
      END IF
      IF TRIM(file_path) = "" THEN
          file_path = Form.Files.Item(File.Name).FileName
      END IF

      file_path = Replace(file_path, "\", "/")
      file_path = Replace(file_path, "../", "")
      file_path = Replace(file_path, "/..", "")

      internal_folder = ""
      file_name = fso.GetBaseName(file_path)&"."&extension

      IF InStr(file_path, "/") > 0 THEN
          internal_folder = Left(file_path, InStrRev(file_path, "/") - 1)
          file_name = fso.GetBaseName(Mid(file_path, InStrRev(file_path, "/") + 1))&"."&extension
      END IF

      IF TRIM(file_name) = "."&extension THEN
          file_name = "tmp_"&Form.UploadID&"."&extension
      END IF

      Form.Files.Item(File.Name).FileName = file_name

      relativeTargetPath = parent_folder
      IF internal_folder <> "" THEN
          relativeTargetPath = relativeTargetPath & "/" & internal_folder
      END IF
      relativeTargetPath = relativeTargetPath & "/" & Form.Files.Item(File.Name).FileName

      parent_folder = server.MapPath("\")&"\"&parent_folder

	  If Not fso.FolderExists(parent_folder) Then
          Err.Raise vbObjectError + 1000, "uploadFile", "Parent folder does not exist"
	  End If

      target_folder = parent_folder
      IF internal_folder <> "" THEN
          target_folder = target_folder & "\" & Replace(internal_folder, "/", "\")
          IF Not fso.FolderExists(target_folder) Then
              CreateFolder target_folder
          END IF
      END IF

    %>
    <% Form.Files.Save target_folder %>
    <% IF Err.Number<>0 THEN %>
	    ,{
	    "success": false,
	    "statusMessage": "Error: <%= REPLACE(Err.Description, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "") %><% IF session("user_id")=1 THEN response.write " \n\n"&sSQL %>"
	    }
    <% ELSE 
		response.AddHeader "File-Name", relativeTargetPath
		%>,{
        "uploadId": "<%= Form.UploadID %>",
        "sourceId": "<%= File.Name %>",
				"file": "<%= relativeTargetPath %>",
				"fileExtension": "<%= extension %>",
				"originalFile": "<%= REPLACE(File.FilePath, "\", "\\") %>",
				"fileName":"<%= File.FileName %>",
				"parentFolder":"<%= REPLACE(parent_folder, "\", "\\") %>",
				"status":"success"
			}
	<%  END IF
    NEXT %>
		],
		"statusMessage":"<%= Form.Files.Count %> file(s) uploaded to <%= REPLACE(Request.ServerVariables("HTTP_HOST"), "\", "\\") %> (<%= relativeTargetPath %>)"
		}
<%
	ElseIf Form.State > 10 then
	  Const fsSizeLimit = &HD %>
			<script language="JavaScript">
                var resultObject = new Object();
                resultObject.status = "error";
                resultObject.statusMessage = "<% Select case Form.State
			case fsSizeLimit: %> Source form size(<%= Form.TotalBytes %> B) exceeds form limit(<%= Form.SizeLimit %> B) <% case else %> Some form error.<% end Select %> ";
                alert(resultObject.statusMessage)
                window.close();
            </script>	
	<%	response.end
	End If'Form.State = 0 then %>
<% END IF %>