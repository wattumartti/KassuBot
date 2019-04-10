using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using DSharpPlus.Entities;
using Newtonsoft.Json.Linq;
using RestSharp;

namespace TestBot
{
    class Utils
    {
        private const string authorizationToken = "AUTH_TOKEN";
        private const string imgurClientID = "IMGUR_CLIENT_ID";

        internal static DiscordEmbedBuilder cachedTrackEmbed = null;
        internal static string cachedTrack = null;
        internal static string cachedArtist = null;

        /// <summary>
        /// Finds info of the given track and caches the name and artist
        /// </summary>
        /// <param name="trackName"></param>
        /// <param name="artistName"></param>
        /// <returns></returns>
        internal static async Task<JObject> FindAndCacheCurrentTrack(string trackName, string artistName)
        {
            if (string.IsNullOrEmpty(trackName) || string.IsNullOrEmpty(artistName))
            {
                return null;
            }

            // Save searched track and artist
            cachedTrack = trackName;
            cachedArtist = artistName;

            // Create the POST request with parameters
            RestClient tokenClient = new RestClient("https://accounts.spotify.com/api/token");
            RestRequest tokenRequest = new RestRequest(Method.POST);
            tokenRequest.AddParameter("grant_type", "client_credentials");
            tokenRequest.AddHeader("Authorization", authorizationToken);

            // Wait and parse response
            IRestResponse tokenResponse = await tokenClient.ExecuteTaskAsync(tokenRequest);
            string tokenContent = tokenResponse.Content;
            JObject json = JObject.Parse(tokenContent);
            string token = json.GetValue("access_token").ToString();

            // Request search results
            RestClient apiClient = new RestClient("https://api.spotify.com/v1" + "/search?q=track:" + cachedTrack + "%20artist:" + cachedArtist + "&type=track");
            RestRequest searchRequest = new RestRequest(Method.GET);
            searchRequest.AddHeader("Authorization", "Bearer " + token);

            IRestResponse searchResponse = await apiClient.ExecuteTaskAsync(searchRequest);
            string searchContent = searchResponse.Content;
            JObject searchJson = JObject.Parse(searchContent);

            string tracksContent = searchJson.GetValue("tracks").ToString();
            JObject tracksJson = JObject.Parse(tracksContent);

            string foundItemString = tracksJson.GetValue("items").ToArray()[0].ToString();

            JObject foundItemJson = JObject.Parse(foundItemString);
            string foundTrackUri = foundItemJson.GetValue("uri").ToString();
            string[] trackSplit = foundTrackUri.Split(':');
            string trackId = trackSplit[trackSplit.Length - 1];

            // Request track with SpotifyID
            RestClient trackClient = new RestClient("https://api.spotify.com/v1" + "/tracks/" + trackId);
            RestRequest getRequest = new RestRequest(Method.GET);
            getRequest.AddHeader("Authorization", "Bearer " + token);

            IRestResponse getResponse = await trackClient.ExecuteTaskAsync(getRequest);
            string getContent = getResponse.Content;
            JObject responseTrackJson = JObject.Parse(getContent);

            return responseTrackJson;
        }

        /// <summary>
        /// Builds a DiscordEmbed from the given data and caches it
        /// </summary>
        /// <param name="searchData"></param>
        /// <returns></returns>
        internal static DiscordEmbed BuildCachedTrackEmbed(JObject searchData)
        {
            if (searchData == null)
            {
                return null;
            }

            JObject albumObject = JObject.Parse(searchData.GetValue("album").ToString());
            string albumName = albumObject.GetValue("name").ToString();
            string albumImageUrl = albumObject.GetValue("images").ToArray()[0]["url"].ToString();

            List<string> artistNames = new List<string>();

            List<JToken> artistTokenList = searchData["artists"].ToList();

            if (artistTokenList.Count > 0)
            {
                for (int i = 0; i < artistTokenList.Count; i++)
                {
                    artistNames.Add(artistTokenList[i]["name"].ToString());
                }
            }

            string foundTrackName = searchData.GetValue("name").ToString();
            string trackUrl = searchData.GetValue("external_urls")["spotify"].ToString();
            int trackDurationMs = (int)searchData.GetValue("duration_ms");
            int trackDurationMinutes = (trackDurationMs / 1000) / 60;
            int trackDurationSeconds = (trackDurationMs / 1000) % 60;

            string artists = "";

            for (int i = 0; i < artistNames.Count; i++)
            {
                artists += i == 0 ? artistNames[i] : ", " + artistNames[i];
            }

            // Save the embed for use when searching the same track again
            cachedTrackEmbed = new DiscordEmbedBuilder()
                .WithTitle(foundTrackName)
                .WithThumbnailUrl(albumImageUrl)
                .WithAuthor("Currently playing: ")
                .AddField("Artists: ", artists, true)
                .AddField("Album: ", albumName, true)
                .AddField("Duration: ", trackDurationMinutes.ToString() + ":" + trackDurationSeconds.ToString("D2"))
                .WithUrl(trackUrl)
                .WithTimestamp(DateTimeOffset.Now);

            return cachedTrackEmbed;
        }

        /// <summary>
        /// Gets the svg image from an url, converts it to a png and uploads it to imgur. 
        /// Must be used with challonge's API since the image is in svg format and neither Discord or Imgur support svg.
        /// </summary>
        /// <param name="svgUrl"></param>
        /// <param name="svgPath"></param>
        /// <param name="pngPath"></param>
        /// <returns></returns>
        internal static string ConvertImageAndUploadToImgur(string svgUrl, string svgPath, string pngPath)
        {
            using (WebClient wb = new WebClient())
            {
                string source = wb.DownloadString(svgUrl);
                File.WriteAllText(svgPath, source);

                string inkscapeArgs = string.Format(@"-f ""{0}"" -e ""{1}""", svgPath, pngPath);

                Process inkscape = Process.Start(
                  new ProcessStartInfo("C:\\Program Files\\Inkscape\\inkscape.exe", inkscapeArgs));

                inkscape.WaitForExit(3000);
            }

            RestClient imgurClient = new RestClient("https://api.imgur.com/3/image");
            RestRequest imgurRequest = new RestRequest(Method.POST);
            imgurRequest.AddHeader("Authorization", imgurClientID);
            using (Image image = Image.FromFile(pngPath))
            {
                using (MemoryStream m = new MemoryStream())
                {
                    image.Save(m, image.RawFormat);
                    byte[] imageBytes = m.ToArray();

                    // Convert byte[] to Base64 String
                    string base64String = Convert.ToBase64String(imageBytes);
                    imgurRequest.AddParameter("image", base64String);
                }
            }

            IRestResponse imgurResponse = imgurClient.Execute(imgurRequest);
            string imgurContent = imgurResponse.Content;
            JObject imgurJson = JObject.Parse(imgurContent);
            string data = imgurJson.GetValue("data").ToString();
            JObject imgurPayload = JObject.Parse(data);
            string imgurLink = imgurPayload.GetValue("link").ToString();

            return imgurLink;
        }
    }
}
