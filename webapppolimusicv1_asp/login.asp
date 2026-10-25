<% Option Explicit %>
<!--#include file="includes/db_connection.asp" -->
<%
'================================================================
' login.asp - User authentication page
' Validates credentials against TBL_USER and creates the session
'================================================================
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.Buffer = True

Dim strUsername, strPassword
Dim strSQL, rs
Dim errorMsg, photoPath

strUsername = ""
strPassword = ""
errorMsg = ""

' If already logged in, go straight to the dashboard
If IsAuthenticated() Then
    Response.Redirect "index.asp"
    Response.End
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    strUsername = Trim(Request.Form("username"))
    strPassword = Trim(Request.Form("password"))

    If strUsername = "" Or strPassword = "" Then
        errorMsg = "Please enter both username and password."
    Else
        ' Validate in SQL Server so both plain-text and SHA1-hashed
        ' passwords work without extra COM components.
        strSQL = "SELECT ID_USER, USERNAME, USER_TYPE, USER_PHOTO, EMAIL " & _
                 "FROM TBL_USER " & _
                 "WHERE USERNAME = '" & SafeSQL(strUsername) & "' " & _
                 "AND (PASSWORD = '" & SafeSQL(strPassword) & "' " & _
                 "  OR PASSWORD = CONVERT(varchar(40), HASHBYTES('SHA1', '" & SafeSQL(strPassword) & "'), 2))"

        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.CursorType = 0   ' adOpenForwardOnly
        rs.LockType = 1     ' adLockReadOnly
        rs.Open strSQL, conn

        If Not rs.EOF Then
            ' --- Credentials are valid: build the session ---
            Session("userId")   = CStr(rs("ID_USER").Value)
            Session("username") = rs("USERNAME").Value
            Session("userType") = CStr(rs("USER_TYPE").Value)
            Session("email")    = rs("EMAIL").Value

            ' Keep the raw DB value in the session; the pages resolve it
            ' to a real file with ResolveImage() when they render it.
            If IsNull(rs("USER_PHOTO").Value) Then
                Session("userPhoto") = ""
            Else
                Session("userPhoto") = rs("USER_PHOTO").Value
            End If

            rs.Close
            Set rs = Nothing
            conn.Close
            Set conn = Nothing

            Response.Redirect "index.asp"
            Response.End
        Else
            ' Do not reveal whether the username or the password was wrong
            errorMsg = "Incorrect username or password. Please try again."
            rs.Close
            Set rs = Nothing
        End If
    End If
End If

conn.Close
Set conn = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="PoliMusic - Sign in to your music account">
    <title>PoliMusic - Sign In</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="css/styles.css">
</head>
<body class="login-body">
    <a href="#main-content" class="skip-link">Skip to main content</a>

    <main class="login-container" id="main-content">
        <div class="login-logo">
            <svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                <rect x="8" y="8" width="48" height="48" rx="12" fill="url(#grad-login)"/>
                <path d="M24 20v24l5-4V24l-5-4zm12 0v24l5-4V24l-5-4z" fill="#1a1b22"/>
                <circle cx="44" cy="40" r="8" fill="#1a1b22"/>
                <circle cx="44" cy="40" r="4" fill="#FF7847"/>
                <defs>
                    <linearGradient id="grad-login" x1="0" y1="0" x2="64" y2="64">
                        <stop offset="0%" stop-color="#FF7847"/>
                        <stop offset="100%" stop-color="#7C3AED"/>
                    </linearGradient>
                </defs>
            </svg>
            <h1>PoliMusic</h1>
        </div>

        <p class="login-subtitle">Sign in to your music world</p>

        <% If errorMsg <> "" Then %>
            <div class="alert alert-error" role="alert">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="12" y1="8" x2="12" y2="12"/>
                    <line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                <span><%= Server.HTMLEncode(errorMsg) %></span>
            </div>
        <% End If %>

        <form method="post" action="login.asp" novalidate>
            <div class="form-group">
                <label for="username">Username</label>
                <input type="text" id="username" name="username"
                       class="form-input"
                       placeholder="Enter your username"
                       value="<%= Server.HTMLEncode(strUsername) %>"
                       required aria-required="true"
                       autocomplete="username"
                       maxlength="50"
                       autofocus>
            </div>

            <div class="form-group">
                <label for="password">Password</label>
                <input type="password" id="password" name="password"
                       class="form-input"
                       placeholder="Enter your password"
                       required aria-required="true"
                       autocomplete="current-password"
                       maxlength="50">
            </div>

            <button type="submit" class="btn btn-primary btn-block">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                    <path d="M15 3h4a2 2 0 012 2v14a2 2 0 01-2 2h-4M10 17l5-5-5-5M15 12H3"/>
                </svg>
                Sign In
            </button>
        </form>

        <div class="login-hint">
            <strong>Demo accounts:</strong><br>
            userNormal / 12345678 &mdash; normal role<br>
            userAdmin / 12345678 &mdash; admin role
        </div>
    </main>
</body>
</html>
