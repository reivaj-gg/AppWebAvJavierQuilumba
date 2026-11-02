using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                LoadAuthenticatedUser();
                BindSongList();
            }
        }

        /// <summary>
        /// Shows the name and the photo of the authenticated user in the top bar.
        /// </summary>
        private void LoadAuthenticatedUser()
        {
            Entities.User user = Session[Util.Constants.USER] as Entities.User;

            string strUserName = (user != null) ? user.UserName : Page.User.Identity.Name;
            lblUserName.Text = Server.HtmlEncode(strUserName);
            lblUserName.ToolTip = strUserName;

            string strPhoto = (user != null && !string.IsNullOrEmpty(user.Photo))
                ? user.Photo
                : "~/UserImages/default.jpg";
            imgUserPhoto.ImageUrl = ResolveUrl(strPhoto);
            imgUserPhoto.AlternateText = strUserName;
        }

        /// <summary>
        /// Fills the GridView with the songs returned by the business layer and
        /// updates the summary cards (total songs, total plays, most played).
        /// </summary>
        private void BindSongList()
        {
            try
            {
                DataTable dtSongs = new BLL.Song(strConnString).Read();
                gvSongs.DataSource = dtSongs;
                gvSongs.DataBind();
                lblSongCount.Text = (dtSongs.Rows.Count == 1)
                    ? "1 canción"
                    : dtSongs.Rows.Count + " canciones";
                lblMessage.Text = string.Empty;

                ComputeStats(dtSongs);
            }
            catch (Exception ex)
            {
                gvSongs.DataSource = null;
                gvSongs.DataBind();
                lblSongCount.Text = "0 canciones";
                lblTotalSongs.Text = "0";
                lblTotalPlays.Text = "0";
                lblTopSong.Text = "—";
                lblMessage.Text = "No se pudo cargar la lista de canciones: " + Server.HtmlEncode(ex.Message);
            }
        }

        /// <summary>
        /// Derives the dashboard numbers from the song table and writes them to the
        /// summary labels: total songs, total reproductions, and the most played track.
        /// </summary>
        private void ComputeStats(DataTable dtSongs)
        {
            int totalSongs = dtSongs.Rows.Count;
            int totalPlays = 0;
            string topSong = "—";
            int topPlays = -1;

            foreach (DataRow row in dtSongs.Rows)
            {
                int plays = Convert.ToInt32(row["PLAYS"]);
                totalPlays += plays;
                if (plays > topPlays)
                {
                    topPlays = plays;
                    topSong = Convert.ToString(row["SONG_NAME"]);
                }
            }

            lblTotalSongs.Text = totalSongs.ToString();
            lblTotalPlays.Text = totalPlays.ToString("N0", System.Globalization.CultureInfo.GetCultureInfo("es-EC"));
            lblTopSong.Text = string.IsNullOrEmpty(topSong) ? "—" : Server.HtmlEncode(topSong);
        }

        protected void lnkLogout_Click(object sender, EventArgs e)
        {
            Session.Clear();
            Session.Abandon();
            FormsAuthentication.SignOut();
            Response.Redirect("Login.aspx", false);
            Context.ApplicationInstance.CompleteRequest();
        }
    }

}
