<%
'================================================================
' logout.asp - Destroy session and redirect to login
'================================================================
Option Explicit
Response.CodePage = 65001
Response.Charset = "UTF-8"

' --- Abandon the session completely ---
Session.Abandon()

' --- Redirect to login page ---
Response.Redirect "login.asp"
Response.End
%>
