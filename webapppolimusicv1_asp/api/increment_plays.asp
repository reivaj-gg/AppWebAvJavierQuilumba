<% Option Explicit %>
<%
'================================================================
' api/increment_plays.asp - AJAX endpoint to increment song plays
' Called via POST from the JavaScript audio player.
' The connection is opened here (instead of using an #include with a
' parent path) because IIS disables parent paths by default.
'================================================================
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.ContentType = "application/json"
Response.Buffer = True

Const DB_SERVER = "localhost"
Const DB_NAME   = "BDD_PoliMusic"
Const DB_USER   = "usr_polimusic"
Const DB_PASS   = "Politecnica1"

Dim songId, strSQL, result, connApi, rsApi, newPlays

result = "{""success"":false,""message"":""Invalid request""}"

' Only logged-in users may increment the counter
If Len(Session("userId") & "") = 0 Then
    Response.Write "{""success"":false,""message"":""Not authenticated""}"
    Response.End
End If

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Write "{""success"":false,""message"":""POST required""}"
    Response.End
End If

songId = Trim(Request.Form("songId"))

If songId = "" Or Not IsNumeric(songId) Then
    Response.Write "{""success"":false,""message"":""Invalid song id""}"
    Response.End
End If

Set connApi = Server.CreateObject("ADODB.Connection")
connApi.Open "Provider=SQLOLEDB.1;Data Source=" & DB_SERVER & _
             ";Initial Catalog=" & DB_NAME & _
             ";User ID=" & DB_USER & _
             ";Password=" & DB_PASS & ";"

' Increment and return the new value in one round trip
strSQL = "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 " & _
         "WHERE ID_SONG = " & CLng(songId) & "; " & _
         "SELECT PLAYS FROM TBL_SONG WHERE ID_SONG = " & CLng(songId)

Set rsApi = connApi.Execute(strSQL)

Do While rsApi.EOF = False
    If IsNull(rsApi("PLAYS").Value) Then
        newPlays = 0
    Else
        newPlays = CLng(rsApi("PLAYS").Value)
    End If
    rsApi.MoveNext
Loop

result = "{""success"":true,""plays"":" & newPlays & "}"

rsApi.Close
Set rsApi = Nothing
connApi.Close
Set connApi = Nothing

Response.Write result
Response.End
%>
