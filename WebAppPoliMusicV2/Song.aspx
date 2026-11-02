<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>PoliMusic | Lista de canciones</title>

    <link rel="preconnect" href="https://fonts.googleapis.com" />
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin="anonymous" />
    <link rel="stylesheet"
        href="https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&amp;family=Plus+Jakarta+Sans:wght@400;500;600;700;800&amp;display=swap" />
    <link rel="stylesheet" href="Content/Song.css" />
</head>
<body>
    <form id="formSong" runat="server">

        <!-- Fondo dinámico -->
        <div class="aurora" aria-hidden="true">
            <span class="blob-1"></span>
            <span class="blob-2"></span>
            <span class="blob-3"></span>
        </div>

        <header class="topbar">
            <a class="brand" href="Song.aspx">
                <span class="brand-mark" aria-hidden="true">
                    <span class="bar"></span>
                    <span class="bar"></span>
                    <span class="bar"></span>
                    <span class="bar"></span>
                </span>
                <span class="brand-name">Poli<em>Music</em></span>
            </a>

            <div class="user-area">
                <div class="user-chip">
                    <asp:Image ID="imgUserPhoto" runat="server" CssClass="user-avatar" />
                    <span class="user-text">
                        <span class="user-caption">Sesión iniciada</span>
                        <asp:Label ID="lblUserName" runat="server" CssClass="user-name"></asp:Label>
                    </span>
                </div>
                <asp:LinkButton ID="lnkLogout" runat="server" CssClass="logout-link" CausesValidation="false"
                    OnClick="lnkLogout_Click">Cerrar sesión</asp:LinkButton>
            </div>
        </header>

        <main class="page">

            <section class="hero">
                <div class="hero-copy">
                    <span class="eyebrow">Tu biblioteca</span>
                    <h1>Toda tu música,<br /><span class="hl">en un solo lugar</span></h1>
                    <p class="hero-sub">
                        Cada pista de la tabla <b>TBL_SONG</b> lista para sonar. Pulsa play y
                        deja que la biblioteca de <b>PoliMusic</b> haga el resto.
                    </p>
                </div>
                <div class="hero-vinyl" aria-hidden="true"></div>
            </section>

            <section class="stats">
                <article class="stat is-accent">
                    <span class="stat-label">Canciones</span>
                    <asp:Label ID="lblTotalSongs" runat="server" CssClass="stat-value"></asp:Label>
                </article>
                <article class="stat is-mint">
                    <span class="stat-label">Reproducciones</span>
                    <asp:Label ID="lblTotalPlays" runat="server" CssClass="stat-value"></asp:Label>
                </article>
                <article class="stat">
                    <span class="stat-label">Más escuchada</span>
                    <asp:Label ID="lblTopSong" runat="server" CssClass="stat-value is-text"></asp:Label>
                </article>
            </section>

            <section class="panel">
                <div class="panel-head">
                    <div>
                        <h2>Lista de canciones</h2>
                        <p class="panel-sub">Reproduce cualquier pista sin salir de la página</p>
                    </div>
                    <asp:Label ID="lblSongCount" runat="server" CssClass="pill"></asp:Label>
                </div>

                <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="False" DataKeyNames="ID_SONG"
                    CssClass="song-table" GridLines="None" UseAccessibleHeader="true"
                    EmptyDataText="Todavía no hay canciones registradas en TBL_SONG.">
                    <Columns>
                        <asp:BoundField DataField="ID_SONG" HeaderText="ID"
                            HeaderStyle-CssClass="col-id" ItemStyle-CssClass="col-id" />
                        <asp:BoundField DataField="SONG_NAME" HeaderText="Canción"
                            HeaderStyle-CssClass="col-name" ItemStyle-CssClass="col-name" />
                        <asp:BoundField DataField="SONG_PATH" HeaderText="Ruta"
                            HeaderStyle-CssClass="col-path" ItemStyle-CssClass="col-path" />
                        <asp:TemplateField HeaderText="Reproducciones"
                            HeaderStyle-CssClass="col-plays" ItemStyle-CssClass="col-plays">
                            <ItemTemplate>
                                <span class="plays-badge">
                                    <svg viewBox="0 0 24 24" width="15" height="15" aria-hidden="true">
                                        <path fill="currentColor" d="M12 3a9 9 0 0 0-9 9v6a3 3 0 0 0 3 3h2v-8H5v-1a7 7 0 0 1 14 0v1h-3v8h2a3 3 0 0 0 3-3v-6a9 9 0 0 0-9-9z" />
                                    </svg>
                                    <%# Eval("PLAYS") %>
                                </span>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Reproducir"
                            HeaderStyle-CssClass="col-player" ItemStyle-CssClass="col-player">
                            <ItemTemplate>
                                <audio class="song-audio" controls preload="none"
                                    src='<%# ResolveUrl(Convert.ToString(Eval("SONG_PATH"))) %>'></audio>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>

                <asp:Label ID="lblMessage" runat="server" CssClass="message"></asp:Label>
            </section>

            <footer class="foot">
                <span><strong>PoliMusic</strong> · Aplicaciones Web Avanzadas</span>
                <span>EPN · BDD_PoliMusic</span>
            </footer>

        </main>

    </form>
</body>
</html>
