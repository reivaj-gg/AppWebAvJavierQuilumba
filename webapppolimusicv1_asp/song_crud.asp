<% Option Explicit %>
<!--#include file="includes/db_connection.asp" -->
<%
'================================================================
' song_crud.asp - Admin CRUD for songs (Create, Read, Update, Delete)
' Only accessible to admin users (USER_TYPE = 1)
'================================================================
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.Buffer = True

' --- Require admin authentication ---
Call RequireAdmin()

Dim strAction, strMsg, strMsgType
Dim rsSong, rsAllSongs
Dim strSQL

strAction = Request("action")
strMsg = ""
strMsgType = ""

' ================================================================
' PROCESS FORM SUBMISSIONS
' ================================================================
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim postAction
    postAction = Request.Form("action")

    Select Case postAction
        Case "create"
            Dim newName, newPath, newPlays
            newName  = Trim(Request.Form("songName"))
            newPath  = Trim(Request.Form("songPath"))
            newPlays = Trim(Request.Form("songPlays"))
            If newPlays = "" Then newPlays = "0"

            If newName = "" Or newPath = "" Then
                strMsg = "Song name and path are required."
                strMsgType = "error"
            Else
                strSQL = "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) " & _
                         "VALUES ('" & SafeSQL(newName) & "', '" & SafeSQL(newPath) & "', " & SafeSQL(newPlays) & ")"
                On Error Resume Next
                conn.Execute strSQL
                If Err.Number <> 0 Then
                    strMsg = "Error creating song: " & Err.Description
                    strMsgType = "error"
                Else
                    strMsg = "Song '" & newName & "' created successfully."
                    strMsgType = "success"
                End If
                On Error GoTo 0
            End If

        Case "update"
            Dim editId, editName, editPath, editPlays
            editId    = Request.Form("songId")
            editName  = Trim(Request.Form("songName"))
            editPath  = Trim(Request.Form("songPath"))
            editPlays = Trim(Request.Form("songPlays"))
            If editPlays = "" Then editPlays = "0"

            If editName = "" Or editPath = "" Then
                strMsg = "Song name and path are required."
                strMsgType = "error"
            Else
                strSQL = "UPDATE TBL_SONG SET SONG_NAME = '" & SafeSQL(editName) & "', " & _
                         "SONG_PATH = '" & SafeSQL(editPath) & "', PLAYS = " & SafeSQL(editPlays) & " " & _
                         "WHERE ID_SONG = " & SafeSQL(editId)
                On Error Resume Next
                conn.Execute strSQL
                If Err.Number <> 0 Then
                    strMsg = "Error updating song: " & Err.Description
                    strMsgType = "error"
                Else
                    strMsg = "Song updated successfully."
                    strMsgType = "success"
                End If
                On Error GoTo 0
            End If

        Case "delete"
            Dim delId
            delId = Request.Form("songId")

            ' First delete from TBL_PLAYER (foreign key constraint)
            On Error Resume Next
            conn.Execute "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & SafeSQL(delId)
            conn.Execute "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = " & SafeSQL(delId)
            conn.Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = " & SafeSQL(delId)
            conn.Execute "DELETE FROM TBL_SONG WHERE ID_SONG = " & SafeSQL(delId)
            If Err.Number <> 0 Then
                strMsg = "Error deleting song: " & Err.Description
                strMsgType = "error"
            Else
                strMsg = "Song deleted successfully."
                strMsgType = "success"
            End If
            On Error GoTo 0
    End Select
End If

' ================================================================
' LOAD EDIT DATA (if action=edit)
' ================================================================
Dim editSongId, editSongName, editSongPath, editSongPlays
Dim isEditing
isEditing = False

If strAction = "edit" And Request.QueryString("id") <> "" Then
    editSongId = Request.QueryString("id")
    strSQL = "SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG WHERE ID_SONG = " & SafeSQL(editSongId)
    Set rsSong = Server.CreateObject("ADODB.Recordset")
    rsSong.CursorType = 0
    rsSong.LockType = 1
    rsSong.Open strSQL, conn
    If Not rsSong.EOF Then
        editSongId   = rsSong("ID_SONG").Value
        editSongName = rsSong("SONG_NAME").Value
        editSongPath = rsSong("SONG_PATH").Value
        If IsNull(rsSong("PLAYS").Value) Then
            editSongPlays = 0
        Else
            editSongPlays = CLng(rsSong("PLAYS").Value)
        End If
        isEditing = True
    End If
    rsSong.Close
    Set rsSong = Nothing
End If

' ================================================================
' LOAD ALL SONGS FOR TABLE (artists aggregated, one row per song)
' ================================================================
strSQL = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
         "ISNULL(STUFF((SELECT ', ' + a2.ARTIST_NAME + ISNULL(' ' + a2.ARTIST_LASTNAME, '') " & _
         "              FROM TBL_PLAYER p2 " & _
         "              INNER JOIN TBL_ARTIST a2 ON p2.ID_ARTIST = a2.ID_ARTIST " & _
         "              WHERE p2.ID_SONG = s.ID_SONG " & _
         "              FOR XML PATH('')), 1, 2, ''), '') AS ARTISTS " & _
         "FROM TBL_SONG s " & _
         "ORDER BY s.ID_SONG"

Set rsAllSongs = Server.CreateObject("ADODB.Recordset")
rsAllSongs.CursorType = 3    ' adOpenStatic so RecordCount is reliable
rsAllSongs.LockType = 1
rsAllSongs.Open strSQL, conn

Dim userPhotoPath
userPhotoPath = ResolveImage(Session("userPhoto"), "UserImages", "default.jpg")
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="PoliMusic - Admin Song Management">
    <title>PoliMusic - Manage Songs</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="css/styles.css">
</head>
<body>
    <a href="#main-content" class="skip-link">Skip to main content</a>

    <div class="app-layout">
        <!-- SIDEBAR -->
        <aside class="sidebar" role="complementary" aria-label="Main navigation">
            <div class="sidebar-brand">
                <svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                    <rect x="8" y="8" width="48" height="48" rx="12" fill="url(#grad-crud)"/>
                    <path d="M24 20v24l5-4V24l-5-4zm12 0v24l5-4V24l-5-4z" fill="#1a1b22"/>
                    <circle cx="44" cy="40" r="8" fill="#1a1b22"/>
                    <circle cx="44" cy="40" r="4" fill="#FF7847"/>
                    <defs>
                        <linearGradient id="grad-crud" x1="0" y1="0" x2="64" y2="64">
                            <stop offset="0%" stop-color="#FF7847"/>
                            <stop offset="100%" stop-color="#7C3AED"/>
                        </linearGradient>
                    </defs>
                </svg>
                <h1>PoliMusic</h1>
            </div>

            <nav class="sidebar-nav" aria-label="Sections">
                <div class="sidebar-nav-section">Main</div>
                <a href="index.asp" class="sidebar-link">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <path d="M3 12l9-9 9 9M5 10v10a1 1 0 001 1h3v-6h6v6h3a1 1 0 001-1V10"/>
                    </svg>
                    <span>Home</span>
                </a>
                <div class="sidebar-nav-section">Administration</div>
                <a href="song_crud.asp" class="sidebar-link active" aria-current="page">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <path d="M12 20h9M16.5 3.5a2.12 2.12 0 113 3L7 19l-4 1 1-4L16.5 3.5z"/>
                    </svg>
                    <span>Manage Songs</span>
                    <span class="admin-badge">CRUD</span>
                </a>
            </nav>

            <div class="sidebar-user">
                <div class="sidebar-user-avatar">
                    <img src="<%= Server.HTMLEncode(userPhotoPath) %>"
                         alt="<%= Server.HTMLEncode(Session("username")) %>'s avatar"
                         onerror="this.style.display='none';this.parentElement.innerHTML='<span class=\'avatar-placeholder\'><%= UCase(Left(Session("username"),1)) %></span>'">
                </div>
                <div class="sidebar-user-info">
                    <div class="sidebar-user-name"><%= Server.HTMLEncode(Session("username")) %></div>
                    <div class="sidebar-user-role admin">Admin</div>
                </div>
            </div>
        </aside>

        <!-- MAIN CONTENT -->
        <main class="main-content" id="main-content" role="main">
            <div class="page-header">
                <div class="page-title">
                    <h2>Manage Songs</h2>
                    <p>Create, edit, and delete songs in the database</p>
                </div>
                <div class="page-actions">
                    <a href="index.asp" class="btn btn-secondary">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                            <path d="M3 12l9-9 9 9M5 10v10a1 1 0 001 1h3v-6h6v6h3a1 1 0 001-1V10"/>
                        </svg>
                        Back to Home
                    </a>
                    <a href="logout.asp" class="btn btn-secondary" aria-label="Logout">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                            <path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4M16 17l5-5-5-5M21 12H9"/>
                        </svg>
                        Logout
                    </a>
                </div>
            </div>

            <div class="admin-banner" role="status">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                    <path d="M12 2l3 7h7l-5.5 4 2 7L12 16l-6.5 4 2-7L2 9h7l3-7z"/>
                </svg>
                <span>Administrator mode: Full CRUD operations enabled</span>
            </div>

            <!-- Alert Message -->
            <% If strMsg <> "" Then %>
                <div class="alert alert-<%= strMsgType %>" role="<% If strMsgType = "error" Then %>alert<% Else %>status<% End If %>">
                    <% If strMsgType = "success" Then %>
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                            <path d="M20 6L9 17l-5-5"/>
                        </svg>
                    <% ElseIf strMsgType = "error" Then %>
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                            <circle cx="12" cy="12" r="10"/>
                            <line x1="12" y1="8" x2="12" y2="12"/>
                            <line x1="12" y1="16" x2="12.01" y2="16"/>
                        </svg>
                    <% End If %>
                    <span><%= Server.HTMLEncode(strMsg) %></span>
                </div>
            <% End If %>

            <!-- Create/Edit Form -->
            <section aria-labelledby="form-heading" style="margin-bottom: 40px;">
                <div class="section-header">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <% If isEditing Then %>
                            <path d="M12 20h9M16.5 3.5a2.12 2.12 0 113 3L7 19l-4 1 1-4L16.5 3.5z"/>
                        <% Else %>
                            <path d="M12 5v14M5 12h14"/>
                        <% End If %>
                    </svg>
                    <h4 id="form-heading"><% If isEditing Then %>Edit Song<% Else %>Add New Song<% End If %></h4>
                </div>

                <form method="post" action="song_crud.asp" class="modal" style="position:relative;display:block;max-width:none;">
                    <input type="hidden" name="action" value="<% If isEditing Then %>update<% Else %>create<% End If %>">
                    <% If isEditing Then %>
                        <input type="hidden" name="songId" value="<%= editSongId %>">
                    <% End If %>

                    <div class="modal-body" style="padding-top: 24px;">
                        <div class="form-row">
                            <div class="form-section">
                                <label for="songName" style="display:block;font-size:0.8125rem;font-weight:500;color:var(--text-secondary);margin-bottom:6px;font-family:'JetBrains Mono',monospace;text-transform:uppercase;letter-spacing:0.05em;">Song Name *</label>
                                <input type="text" id="songName" name="songName"
                                       class="form-input-extended"
                                       placeholder="e.g. Adventure"
                                       value="<% If isEditing Then %><%= Server.HTMLEncode(editSongName) %><% End If %>"
                                       required aria-required="true"
                                       maxlength="50">
                            </div>
                            <div class="form-section">
                                <label for="songPlays" style="display:block;font-size:0.8125rem;font-weight:500;color:var(--text-secondary);margin-bottom:6px;font-family:'JetBrains Mono',monospace;text-transform:uppercase;letter-spacing:0.05em;">Initial Plays</label>
                                <input type="number" id="songPlays" name="songPlays"
                                       class="form-input-extended"
                                       placeholder="0"
                                       value="<% If isEditing Then %><%= editSongPlays %><% Else %>0<% End If %>"
                                       min="0"
                                       step="1">
                            </div>
                        </div>
                        <div class="form-section" style="margin-top:16px;">
                            <label for="songPath" style="display:block;font-size:0.8125rem;font-weight:500;color:var(--text-secondary);margin-bottom:6px;font-family:'JetBrains Mono',monospace;text-transform:uppercase;letter-spacing:0.05em;">Song Path *</label>
                            <input type="text" id="songPath" name="songPath"
                                   class="form-input-extended"
                                   placeholder="~/songFiles/bensound-adventure.mp3"
                                   value="<% If isEditing Then %><%= Server.HTMLEncode(editSongPath) %><% End If %>"
                                   required aria-required="true"
                                   maxlength="255">
                            <p style="font-size:0.75rem;color:var(--text-tertiary);margin-top:6px;font-family:'JetBrains Mono',monospace;">
                                Use format: ~/songFiles/your-song.mp3
                            </p>
                        </div>
                    </div>

                    <div class="modal-footer" style="padding-top: 0;">
                        <% If isEditing Then %>
                            <a href="song_crud.asp" class="btn btn-secondary">Cancel Edit</a>
                            <button type="submit" class="btn btn-primary">
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                                    <path d="M20 6L9 17l-5-5"/>
                                </svg>
                                Update Song
                            </button>
                        <% Else %>
                            <button type="reset" class="btn btn-secondary">Clear</button>
                            <button type="submit" class="btn btn-success">
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                                    <path d="M12 5v14M5 12h14"/>
                                </svg>
                                Create Song
                            </button>
                        <% End If %>
                    </div>
                </form>
            </section>

            <!-- Songs Table -->
            <section aria-labelledby="table-heading">
                <div class="section-header">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <path d="M3 3h18v18H3zM3 9h18M3 15h18M9 3v18M15 3v18"/>
                    </svg>
                    <h4 id="table-heading">All Songs</h4>
                    <span class="section-subtitle"><% If Not rsAllSongs.EOF Then %><%= rsAllSongs.RecordCount %> records<% Else %>0 records<% End If %></span>
                </div>

                <% If rsAllSongs.EOF Then %>
                    <div class="empty-state">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true">
                            <path d="M9 18V5l12-2v13"/>
                            <circle cx="6" cy="18" r="3"/>
                            <circle cx="18" cy="16" r="3"/>
                        </svg>
                        <p>No songs in the database yet. Use the form above to add one.</p>
                    </div>
                <% Else %>
                    <table class="data-table" role="table">
                        <thead>
                            <tr>
                                <th scope="col">ID</th>
                                <th scope="col">Song Name</th>
                                <th scope="col">Path</th>
                                <th scope="col">Artist</th>
                                <th scope="col">Plays</th>
                                <th scope="col">Actions</th>
                            </tr>
                        </thead>
                        <tbody>
                            <% Do While Not rsAllSongs.EOF %>
                                <%
                                Dim tblId, tblName, tblPath, tblPlays, tblArtist
                                tblId = rsAllSongs("ID_SONG").Value
                                tblName = rsAllSongs("SONG_NAME").Value
                                tblPath = rsAllSongs("SONG_PATH").Value
                                If IsNull(rsAllSongs("PLAYS").Value) Then
                                    tblPlays = 0
                                Else
                                    tblPlays = CLng(rsAllSongs("PLAYS").Value)
                                End If
                                tblArtist = Trim(rsAllSongs("ARTISTS").Value)
                                If tblArtist = "" Then tblArtist = "--"
                                %>
                                <tr data-song-id="<%= tblId %>" data-song-title="<%= SafeHTML(tblName) %>" data-song-path="<%= SafeHTML(tblPath) %>" data-song-plays="<%= tblPlays %>" data-song-artist="<%= SafeHTML(tblArtist) %>">
                                    <td class="table-id"><%= tblId %></td>
                                    <td><strong><%= SafeHTML(tblName) %></strong></td>
                                    <td style="font-family:'JetBrains Mono',monospace;font-size:0.75rem;color:var(--text-secondary);max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"><%= SafeHTML(tblPath) %></td>
                                    <td><%= SafeHTML(tblArtist) %></td>
                                    <td style="font-family:'JetBrains Mono',monospace;color:var(--accent-success);"><%= tblPlays %></td>
                                    <td>
                                        <div class="table-actions">
                                            <a href="song_crud.asp?action=edit&id=<%= tblId %>" class="btn-icon btn-icon-success" aria-label="Edit <%= SafeHTML(tblName) %>" title="Edit">
                                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14" aria-hidden="true">
                                                    <path d="M12 20h9M16.5 3.5a2.12 2.12 0 113 3L7 19l-4 1 1-4L16.5 3.5z"/>
                                                </svg>
                                            </a>
                                            <button class="btn-icon btn-icon-danger" type="button"
                                                    onclick="deleteSongFromTable(<%= tblId %>, '<%= SafeHTML(tblName) %>')"
                                                    aria-label="Delete <%= SafeHTML(tblName) %>" title="Delete">
                                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14" aria-hidden="true">
                                                    <path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/>
                                                </svg>
                                            </button>
                                        </div>
                                    </td>
                                </tr>
                                <%
                                rsAllSongs.MoveNext
                                Loop
                                %>
                        </tbody>
                    </table>
                <% End If %>
            </section>
        </main>

        <!-- Player placeholder (not needed on admin page but keeps layout consistent) -->
        <div class="player-bar empty" role="region" aria-label="Audio player">
            <div class="player-placeholder">Song management mode - Return to Home to play music</div>
        </div>
    </div>

    <!-- Delete Confirmation Modal -->
    <div class="modal-overlay" id="deleteModal" role="dialog" aria-labelledby="deleteModalTitle" aria-modal="true">
        <div class="modal">
            <div class="modal-header">
                <h3 id="deleteModalTitle">Delete Song</h3>
                <button class="btn-icon" type="button" onclick="closeModal('deleteModal')" aria-label="Close dialog">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                        <path d="M18 6L6 18M6 6l12 12"/>
                    </svg>
                </button>
            </div>
            <div class="modal-body">
                <p>Are you sure you want to delete <strong id="deleteSongName">this song</strong>?</p>
                <p style="color: var(--text-secondary); font-size: 0.8125rem; margin-top: 8px;">
                    This will also remove it from all tracklists, playlists, and player associations. This action cannot be undone.
                </p>
            </div>
            <div class="modal-footer">
                <button class="btn btn-secondary" type="button" onclick="closeModal('deleteModal')">Cancel</button>
                <form method="post" action="song_crud.asp" style="display:inline;">
                    <input type="hidden" name="action" value="delete">
                    <input type="hidden" name="songId" id="deleteSongId">
                    <button type="submit" class="btn btn-danger">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                            <path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/>
                        </svg>
                        Delete Permanently
                    </button>
                </form>
            </div>
        </div>
    </div>

    <script src="js/app.js"></script>
    <script>
        // CRUD-specific delete function that works with table rows
        function deleteSongFromTable(songId, songName) {
            document.getElementById('deleteSongId').value = songId;
            document.getElementById('deleteSongName').textContent = songName;
            openModal('deleteModal');
        }
    </script>
    <%
    rsAllSongs.Close
    Set rsAllSongs = Nothing
    conn.Close
    Set conn = Nothing
    %>
</body>
</html>
