using System;
using System.Collections.Generic;
using System.Data;
using System.Web;

namespace WebAppPoliMusicV2.BLL
{
    /// <summary>
    /// Business Logic Layer for the songs. It validates the information and calls the DAL.
    /// </summary>
    public class Song
    {
        private string strConnectionString;
        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }
        public int Save(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Save(song);
        }
        public DataTable Read()
        {
            return new DAL.Song(strConnectionString).Read();
        }
        public Entities.Song GetById(int id)
        {
            return new DAL.Song(strConnectionString).GetById(id);
        }
        public int Update(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Update(song);
        }
        public int Delete(int id)
        {
            return new DAL.Song(strConnectionString).Delete(id);
        }
        /// <summary>
        /// Validates the song information against the columns of [TBL_SONG].
        /// </summary>
        /// <returns>An empty string when the song is valid, otherwise the error message.</returns>
        public string Validate(Entities.Song song)
        {
            if (song == null)
            {
                return "The song information is required.";
            }
            if (string.IsNullOrEmpty(song.Name))
            {
                return "The song name is required.";
            }
            if (song.Name.Length > 50)
            {
                return "The song name cannot exceed 50 characters.";
            }
            if (string.IsNullOrEmpty(song.Path))
            {
                return "The song path is required.";
            }
            if (song.Path.Length > 255)
            {
                return "The song path cannot exceed 255 characters.";
            }
            if (song.Plays < 0)
            {
                return "The number of plays cannot be negative.";
            }
            return string.Empty;
        }
    }

}
