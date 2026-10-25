<% Option Explicit %>
<!--#include file="includes/db_connection.asp" -->
<%
'================================================================
' index.asp - Main dashboard page
' Welcome message + Song List (HTML table) + audio player
'================================================================
Response.CodePage = 65001
Response.CharSet = "UTF-8"
Response.Buffer = True

' --- Require authentication ---
Call RequireAuth()

Dim rs, strSQL
Dim totalSongs, totalPlays, totalUsers
Dim userPhotoPath, userInitial, userRoleLabel
Dim rsSongs, rsPlaylists

' --- User data coming from the session set in login.asp ---
userPhotoPath = ResolveImage(Session("userPhoto"), "UserImages", "default.jpg")
userInitial = UCase(Left(Session("username"), 1))
If IsAdmin() Then
    userRoleLabel = "Administrator"
Else
    userRoleLabel = "Normal User"
End If

' --- Dashboard counters ---
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorType = 0
rs.LockType = 1

rs.Open "SELECT COUNT(*) AS cnt FROM TBL_SONG", conn
totalSongs = CLng(rs("cnt").Value)
rs.Close

rs.Open "SELECT ISNULL(SUM(PLAYS), 0) AS total FROM TBL_SONG", conn
totalPlays = CLng(rs("total").Value)
rs.Close

rs.Open "SELECT COUNT(*) AS cnt FROM TBL_USER", conn
totalUsers = CLng(rs("cnt").Value)
rs.Close

Set rs = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="PoliMusic - Browse and play your music collection">
    <title>PoliMusic - Home</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="css/styles.css">
</head>
<body>
    <a href="#main-content" class="skip-link">Skip to main content</a>

    <div class="app-layout">
        <!-- ============================ SIDEBAR ============================ -->
        <aside class="sidebar" role="complementary" aria-label="Main navigation">
            <div class="sidebar-brand">
                <svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                    <rect x="8" y="8" width="48" height="48" rx="12" fill="url(#grad-sidebar)"/>
                    <path d="M24 20v24l5-4V24l-5-4zm12 0v24l5-4V24l-5-4z" fill="#1a1b22"/>
                    <circle cx="44" cy="40" r="8" fill="#1a1b22"/>
                    <circle cx="44" cy="40" r="4" fill="#FF7847"/>
                    <defs>
                        <linearGradient id="grad-sidebar" x1="0" y1="0" x2="64" y2="64">
                            <stop offset="0%" stop-color="#FF7847"/>
                            <stop offset="100%" stop-color="#7C3AED"/>
                        </linearGradient>
                    </defs>
                </svg>
                <h1>PoliMusic</h1>
            </div>

            <nav class="sidebar-nav" aria-label="Sections">
                <div class="sidebar-nav-section">Main</div>
                <a href="index.asp" class="sidebar-link active" aria-current="page">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <path d="M3 12l9-9 9 9M5 10v10a1 1 0 001 1h3v-6h6v6h3a1 1 0 001-1V10"/>
                    </svg>
                    <span>Home</span>
                </a>
                <a href="#song-list" class="sidebar-link">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <path d="M9 18V5l12-2v13M9 9l12-2"/>
                        <circle cx="6" cy="18" r="3"/>
                        <circle cx="18" cy="16" r="3"/>
                    </svg>
                    <span>Song List</span>
                </a>
                <a href="#playlists" class="sidebar-link">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <path d="M21 15V6M3 4h14v11M18 5l3-1v10"/>
                        <circle cx="7" cy="18" r="3"/>
                        <circle cx="17" cy="16" r="3"/>
                    </svg>
                    <span>Playlists</span>
                </a>

                <% If IsAdmin() Then %>
                <div class="sidebar-nav-section">Administration</div>
                <a href="song_crud.asp" class="sidebar-link">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <path d="M12 20h9M16.5 3.5a2.12 2.12 0 113 3L7 19l-4 1 1-4L16.5 3.5z"/>
                    </svg>
                    <span>Manage Songs</span>
                    <span class="admin-badge">CRUD</span>
                </a>
                <% End If %>
            </nav>

            <div class="sidebar-user">
                <div class="sidebar-user-avatar">
                    <img src="<%= Server.HTMLEncode(userPhotoPath) %>"
                         alt="Avatar of <%= Server.HTMLEncode(Session("username")) %>"
                         onerror="this.style.display='none';this.parentElement.innerHTML='<span class=&quot;avatar-placeholder&quot;><%= userInitial %></span>'">
                </div>
                <div class="sidebar-user-info">
                    <div class="sidebar-user-name"><%= Server.HTMLEncode(Session("username")) %></div>
                    <div class="sidebar-user-role<% If IsAdmin() Then %> admin<% End If %>">
                        <% If IsAdmin() Then %>Admin<% Else %>Normal<% End If %>
                    </div>
                </div>
            </div>
        </aside>

        <!-- ============================ MAIN ============================ -->
        <main class="main-content" id="main-content" role="main">
            <div class="page-header">
                <div class="page-title">
                    <h2>Welcome to PoliMusic</h2>
                    <p>Your personal music streaming platform</p>
                </div>
                <div class="page-actions">
                    <a href="logout.asp" class="btn btn-secondary" aria-label="Close session and logout">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                            <path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4M16 17l5-5-5-5M21 12H9"/>
                        </svg>
                        Logout
                    </a>
                </div>
            </div>

            <!-- ===== Welcome message (reads the session variable) ===== -->
            <section class="welcome-banner" aria-label="Welcome message">
                <div class="welcome-icon welcome-icon-photo">
                    <img src="<%= Server.HTMLEncode(userPhotoPath) %>"
                         alt="Profile picture of <%= Server.HTMLEncode(Session("username")) %>"
                         onerror="this.onerror=null;this.src='SongFilesAndImages_resources/UserImages/default.jpg'">
                </div>
                <div class="welcome-text">
                    <h3>Welcome, <span class="username-highlight"><%= Server.HTMLEncode(Session("username")) %></span>!</h3>
                    <p>
                        You are signed in as <strong><%= userRoleLabel %></strong>.
                        <% If IsAdmin() Then %>
                            You have full access to create, read, update and delete songs.
                        <% Else %>
                            Explore the catalogue and press play to listen.
                        <% End If %>
                    </p>
                </div>
            </section>

            <% If IsAdmin() Then %>
            <div class="admin-banner" role="status">
                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                    <path d="M12 2l3 7h7l-5.5 4 2 7L12 16l-6.5 4 2-7L2 9h7l3-7z"/>
                </svg>
                <span>Administrator session &mdash; CRUD operations are enabled for this account.</span>
            </div>
            <% End If %>

            <!-- ===== Counters ===== -->
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-card-icon orange">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <path d="M9 18V5l12-2v13"/>
                            <circle cx="6" cy="18" r="3"/>
                            <circle cx="18" cy="16" r="3"/>
                        </svg>
                    </div>
                    <div class="stat-card-info">
                        <div class="stat-card-value"><%= totalSongs %></div>
                        <div class="stat-card-label">Songs</div>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-icon green">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <polygon points="5 3 19 12 5 21 5 3"/>
                        </svg>
                    </div>
                    <div class="stat-card-info">
                        <div class="stat-card-value"><%= totalPlays %></div>
                        <div class="stat-card-label">Total plays</div>
                    </div>
                </div>
                <div class="stat-card">
                    <div class="stat-card-icon purple">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                            <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/>
                            <circle cx="9" cy="7" r="4"/>
                            <path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75"/>
                        </svg>
                    </div>
                    <div class="stat-card-info">
                        <div class="stat-card-value"><%= totalUsers %></div>
                        <div class="stat-card-label">Users</div>
                    </div>
                </div>
            </div>

            <!-- ===== SONG LIST (HTML table) ===== -->
            <section id="song-list" aria-labelledby="song-list-heading">
                <div class="section-header">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <path d="M9 18V5l12-2v13"/>
                        <circle cx="6" cy="18" r="3"/>
                        <circle cx="18" cy="16" r="3"/>
                    </svg>
                    <h4 id="song-list-heading">Song List</h4>
                    <span class="section-subtitle"><%= totalSongs %> tracks available</span>
                </div>

                <p class="section-lead">
                    Below is the <strong>Song List</strong> with every track stored in the database.
                    Select a row (or press its play button) to start listening.
                </p>

                <%
                ' One row per song. Artists are aggregated so a song with
                ' several artists does not duplicate the row.
                strSQL = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
                         "ISNULL(STUFF((SELECT ', ' + a2.ARTIST_NAME + ISNULL(' ' + a2.ARTIST_LASTNAME, '') " & _
                         "              FROM TBL_PLAYER p2 " & _
                         "              INNER JOIN TBL_ARTIST a2 ON p2.ID_ARTIST = a2.ID_ARTIST " & _
                         "              WHERE p2.ID_SONG = s.ID_SONG " & _
                         "              FOR XML PATH('')), 1, 2, ''), '') AS ARTISTS " & _
                         "FROM TBL_SONG s " & _
                         "ORDER BY s.ID_SONG"

                Set rsSongs = Server.CreateObject("ADODB.Recordset")
                rsSongs.CursorType = 0
                rsSongs.LockType = 1
                rsSongs.Open strSQL, conn
                %>

                <% If rsSongs.EOF Then %>
                    <div class="empty-state">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true">
                            <path d="M9 18V5l12-2v13"/>
                            <circle cx="6" cy="18" r="3"/>
                            <circle cx="18" cy="16" r="3"/>
                        </svg>
                        <p>No songs available yet.<% If IsAdmin() Then %> <a href="song_crud.asp">Add your first song</a>.<% End If %></p>
                    </div>
                <% Else %>
                    <div class="table-wrap">
                        <table class="song-table" aria-describedby="song-list-heading">
                            <caption class="sr-only">Song list with title, artist, play count and actions</caption>
                            <thead>
                                <tr>
                                    <th scope="col" class="col-num">#</th>
                                    <th scope="col">Title</th>
                                    <th scope="col">Artist</th>
                                    <th scope="col" class="col-plays-head">Plays</th>
                                    <th scope="col" class="col-actions-head">Actions</th>
                                </tr>
                            </thead>
                            <tbody>
                            <%
                            Dim rowNum, songId, songName, songPath, songPlays, artistName
                            rowNum = 0
                            Do While Not rsSongs.EOF
                                rowNum = rowNum + 1
                                songId    = rsSongs("ID_SONG").Value
                                songName  = rsSongs("SONG_NAME").Value
                                songPath  = rsSongs("SONG_PATH").Value
                                songPlays = CLng(rsSongs("PLAYS").Value)
                                artistName = Trim(rsSongs("ARTISTS").Value)
                                If artistName = "" Then artistName = "Unknown Artist"
                            %>
                                <tr class="song-row"
                                    data-song-id="<%= songId %>"
                                    data-song-title="<%= SafeHTML(songName) %>"
                                    data-song-path="<%= SafeHTML(songPath) %>"
                                    data-song-plays="<%= songPlays %>"
                                    data-song-artist="<%= SafeHTML(artistName) %>"
                                    tabindex="0"
                                    aria-label="Play <%= SafeHTML(songName) %> by <%= SafeHTML(artistName) %>">
                                    <td class="col-num">
                                        <span class="row-index"><%= rowNum %></span>
                                        <button class="row-play" type="button" tabindex="-1"
                                                aria-label="Play <%= SafeHTML(songName) %>">
                                            <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                                                <path d="M8 5v14l11-7z"/>
                                            </svg>
                                        </button>
                                    </td>
                                    <td>
                                        <div class="row-title-cell">
                                            <span class="row-cover"><%= UCase(Left(songName, 1)) %></span>
                                            <span class="row-title-text"><%= SafeHTML(songName) %></span>
                                        </div>
                                    </td>
                                    <td class="row-artist"><%= SafeHTML(artistName) %></td>
                                    <td class="row-plays"><span class="plays-count"><%= songPlays %></span></td>
                                    <td>
                                        <div class="row-actions">
                                            <button class="btn-icon btn-icon-play" type="button"
                                                    onclick="event.stopPropagation();window.playRow(<%= songId %>)"
                                                    aria-label="Play <%= SafeHTML(songName) %>" title="Play">
                                                <svg viewBox="0 0 24 24" fill="currentColor" width="13" height="13" aria-hidden="true">
                                                    <path d="M8 5v14l11-7z"/>
                                                </svg>
                                            </button>
                                            <% If IsAdmin() Then %>
                                            <a class="btn-icon btn-icon-success"
                                               href="song_crud.asp?action=edit&id=<%= songId %>"
                                               aria-label="Edit <%= SafeHTML(songName) %>" title="Edit">
                                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="13" height="13" aria-hidden="true">
                                                    <path d="M12 20h9M16.5 3.5a2.12 2.12 0 113 3L7 19l-4 1 1-4L16.5 3.5z"/>
                                                </svg>
                                            </a>
                                            <button class="btn-icon btn-icon-danger" type="button"
                                                    onclick="event.stopPropagation();deleteSong(<%= songId %>)"
                                                    aria-label="Delete <%= SafeHTML(songName) %>" title="Delete">
                                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="13" height="13" aria-hidden="true">
                                                    <path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/>
                                                </svg>
                                            </button>
                                            <% End If %>
                                        </div>
                                    </td>
                                </tr>
                            <%
                                rsSongs.MoveNext
                            Loop
                            %>
                            </tbody>
                        </table>
                    </div>
                <% End If

                rsSongs.Close
                Set rsSongs = Nothing
                %>
            </section>

            <!-- ===== PLAYLISTS ===== -->
            <section id="playlists" aria-labelledby="playlists-heading" style="margin-top: 40px;">
                <div class="section-header">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <path d="M21 15V6M3 4h14v11M18 5l3-1v10"/>
                        <circle cx="7" cy="18" r="3"/>
                        <circle cx="17" cy="16" r="3"/>
                    </svg>
                    <h4 id="playlists-heading">Your Playlists</h4>
                </div>

                <%
                ' The user's own playlists plus every public one (TYPE = 0)
                strSQL = "SELECT TOP 5 ID_PLAYLIST, TITLE, CREATION_DATE, TYPE, PHOTO " & _
                         "FROM TBL_PLAYLIST " & _
                         "WHERE ID_USER = " & SafeSQL(CStr(Session("userId"))) & " OR TYPE = 0 " & _
                         "ORDER BY CREATION_DATE DESC"

                Set rsPlaylists = Server.CreateObject("ADODB.Recordset")
                rsPlaylists.CursorType = 0
                rsPlaylists.LockType = 1
                rsPlaylists.Open strSQL, conn
                %>

                <% If rsPlaylists.EOF Then %>
                    <div class="empty-state">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true">
                            <path d="M21 15V6M3 4h14v11M18 5l3-1v10"/>
                            <circle cx="7" cy="18" r="3"/>
                            <circle cx="17" cy="16" r="3"/>
                        </svg>
                        <p>This account has no playlists yet.</p>
                    </div>
                <% Else %>
                    <div class="playlist-grid" role="list" aria-label="Your playlists">
                        <%
                        Dim plTitle, plPhoto, plType, plTypeBadge, plDate
                        Do While Not rsPlaylists.EOF
                            plTitle = rsPlaylists("TITLE").Value
                            plPhoto = rsPlaylists("PHOTO").Value
                            plType  = CInt(rsPlaylists("TYPE").Value)
                            plDate  = rsPlaylists("CREATION_DATE").Value

                            plPhoto = ResolveImage(plPhoto, "PlaylistImages", "default.jpg")

                            If plType = 1 Then
                                plTypeBadge = "<span class='badge-private'>Private</span>"
                            Else
                                plTypeBadge = "<span class='badge-public'>Public</span>"
                            End If
                        %>
                            <div class="playlist-card" role="listitem" tabindex="0" aria-label="Playlist <%= SafeHTML(plTitle) %>">
                                <div class="playlist-card-cover">
                                    <img src="<%= SafeHTML(plPhoto) %>" alt="Cover of <%= SafeHTML(plTitle) %>" loading="lazy"
                                         onerror="this.src='SongFilesAndImages_resources/PlaylistImages/default.jpg'">
                                </div>
                                <div class="playlist-card-info">
                                    <div class="playlist-card-title"><%= SafeHTML(plTitle) %></div>
                                    <div class="playlist-card-meta">
                                        <%= plTypeBadge %>
                                        <span><%= FormatDateTime(plDate, 2) %></span>
                                    </div>
                                </div>
                            </div>
                        <%
                            rsPlaylists.MoveNext
                        Loop
                        %>
                    </div>
                <% End If

                rsPlaylists.Close
                Set rsPlaylists = Nothing
                %>
            </section>
        </main>

        <!-- ============================ PLAYER ============================ -->
        <div class="player-bar empty" role="region" aria-label="Audio player">
            <div class="player-info">
                <div class="player-cover">
                    <span class="cover-placeholder">--</span>
                </div>
                <div class="player-text">
                    <div class="player-title">No song selected</div>
                    <div class="player-artist">--</div>
                </div>
            </div>

            <div class="player-controls">
                <div class="player-buttons">
                    <button class="player-btn" id="prevBtn" type="button" aria-label="Previous song" title="Previous (Shift+Left)">
                        <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                            <path d="M19 20L9 12l10-8v16zM5 4h2v16H5V4z"/>
                        </svg>
                    </button>
                    <button class="player-btn play-btn" id="playBtn" type="button" aria-label="Play or pause" title="Play / Pause (Space)">
                        <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                            <path d="M8 5v14l11-7z"/>
                        </svg>
                    </button>
                    <button class="player-btn" id="nextBtn" type="button" aria-label="Next song" title="Next (Shift+Right)">
                        <svg viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
                            <path d="M5 4l10 8-10 8V4zM17 4h2v16h-2V4z"/>
                        </svg>
                    </button>
                </div>
                <div class="player-progress">
                    <span class="player-time" id="currentTime">0:00</span>
                    <div class="progress-bar" id="progressBar" role="slider" aria-label="Seek" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0" tabindex="0">
                        <div class="progress-fill" id="progressFill"></div>
                    </div>
                    <span class="player-time" id="duration">0:00</span>
                </div>
            </div>

            <div class="player-volume">
                <button class="player-btn" id="muteBtn" type="button" aria-label="Mute or unmute" title="Mute (M)">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                        <polygon points="11 5 6 9 2 9 2 15 6 15 11 19 11 5" fill="currentColor"/>
                        <path d="M15.54 8.46a5 5 0 010 7.07M19.07 4.93a10 10 0 010 14.14"/>
                    </svg>
                </button>
                <div class="volume-bar" id="volumeBar" role="slider" aria-label="Volume" aria-valuemin="0" aria-valuemax="100" aria-valuenow="70" tabindex="0">
                    <div class="volume-fill" id="volumeFill"></div>
                </div>
            </div>

            <div class="player-placeholder">Select a song from the list above to start playing</div>
        </div>
    </div>

    <% If IsAdmin() Then %>
    <div class="modal-overlay" id="deleteModal" role="dialog" aria-labelledby="deleteModalTitle" aria-modal="true">
        <div class="modal">
            <div class="modal-header">
                <h3 id="deleteModalTitle">Delete song</h3>
                <button class="btn-icon" type="button" onclick="closeModal('deleteModal')" aria-label="Close dialog">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16" aria-hidden="true">
                        <path d="M18 6L6 18M6 6l12 12"/>
                    </svg>
                </button>
            </div>
            <div class="modal-body">
                <p>Are you sure you want to delete <strong id="deleteSongName">this song</strong>?</p>
                <p style="color: var(--text-secondary); font-size: 0.8125rem; margin-top: 8px;">This action cannot be undone.</p>
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
                        Delete
                    </button>
                </form>
            </div>
        </div>
    </div>
    <% End If %>

    <script src="js/app.js"></script>
    <%
    conn.Close
    Set conn = Nothing
    %>
</body>
</html>
