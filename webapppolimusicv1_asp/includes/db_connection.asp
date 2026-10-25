<%
'================================================================
' db_connection.asp - Centralized DB connection for PoliMusic
' ASP Classic + VBScript + ADO + SQL Server
'================================================================

' --- Database configuration ---
Const DB_SERVER = "localhost"
Const DB_NAME   = "BDD_PoliMusic"
Const DB_USER   = "usr_polimusic"
Const DB_PASS   = "Politecnica1"

' --- Connection string (SQL Server OLEDB Provider) ---
Dim CONN_STRING
CONN_STRING = "Provider=SQLOLEDB.1;" & _
              "Data Source=" & DB_SERVER & ";" & _
              "Initial Catalog=" & DB_NAME & ";" & _
              "User ID=" & DB_USER & ";" & _
              "Password=" & DB_PASS & ";" & _
              "Network Library=DBMSSOCN;"

' --- Global connection object ---
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")

' --- Open connection ---
On Error Resume Next
conn.Open CONN_STRING
If Err.Number <> 0 Then
    Response.Status = "500 Internal Server Error"
    Response.Write "<div style='font-family:monospace;padding:2rem;background:#1e1e2e;color:#ff5555;border-radius:8px;margin:2rem;'>"
    Response.Write "<h2>Database Connection Error</h2>"
    Response.Write "<p>" & Err.Description & "</p>"
    Response.Write "<p>Server: " & DB_SERVER & " | DB: " & DB_NAME & " | User: " & DB_USER & "</p>"
    Response.Write "</div>"
    Response.End
End If
On Error GoTo 0

' --- Helper: escape single quotes for SQL ---
Function SafeSQL(ByVal val)
    SafeSQL = Replace(val, "'", "''")
End Function

' --- Helper: sanitize HTML output ---
Function SafeHTML(ByVal val)
    If IsNull(val) Then
        SafeHTML = ""
    Else
        SafeHTML = Server.HTMLEncode(CStr(val))
    End If
End Function

' --- Helper: check if user is logged in ---
Function IsAuthenticated()
    IsAuthenticated = (Session("userId") <> "" And Not IsNull(Session("userId")))
End Function

' --- Helper: check if user is admin ---
Function IsAdmin()
    IsAdmin = (Session("userType") = "1")
End Function

' --- Helper: resolve a DB image path to a real, existing web path ---
' DB values look like "~/userImages/root.png" or "~/playlistImages/Playlist2.png".
' Only the file name is kept, so the folder casing in the DB never matters,
' and a missing file falls back to the folder's default image.
Function ResolveImage(ByVal rawPath, ByVal folderName, ByVal defaultFile)
    Dim fso, fileName, candidate, fallback

    fallback = "SongFilesAndImages_resources/" & folderName & "/" & defaultFile
    ResolveImage = fallback

    If IsNull(rawPath) Then Exit Function
    rawPath = Trim(rawPath & "")
    If rawPath = "" Then Exit Function

    ' Strip everything before the last "/" to get just the file name
    fileName = Mid(rawPath, InStrRev(rawPath, "/") + 1)
    If fileName = "" Then Exit Function

    candidate = "SongFilesAndImages_resources/" & folderName & "/" & fileName

    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Err.Number = 0 Then
        If fso.FileExists(Server.MapPath(candidate)) Then
            ResolveImage = candidate
        End If
        Set fso = Nothing
    End If
    On Error GoTo 0
End Function

' --- Guard: redirect to login if not authenticated ---
Sub RequireAuth()
    If Not IsAuthenticated() Then
        Response.Redirect "login.asp"
        Response.End
    End If
End Sub

' --- Guard: redirect to index if not admin ---
Sub RequireAdmin()
    RequireAuth()
    If Not IsAdmin() Then
        Response.Redirect "index.asp"
        Response.End
    End If
End Sub
%>
