using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using DSharpPlus;
using DSharpPlus.CommandsNext;
using DSharpPlus.CommandsNext.Attributes;
using DSharpPlus.Entities;
using System.Net;
using System.IO;
using System.Web;
using System.Diagnostics;
using Svg;
using System.Xml.Linq;
using System.Drawing;
using RestSharp;

namespace KassuBot.Commands
{
    public class UserCommands
    {
        private static bool isGettingData = false;
        private static string scheduleLink = string.Empty;

        [Command("song")]
        [Description("Displays currently playing song with a spotify link")]
        public async Task ShowCurrentTrack(CommandContext ctx)
        {
            if (Program.fileOperations.isProcessing || isGettingData)
            {
                string error = ctx.User.Mention + " I am processing another operation, please wait.";
                await ctx.RespondAsync(error);
                return;
            }

            DiscordUser spotifyUser = ctx.Guild.Members.First(x => x.Username == Program.spotifyUserName);

            // Return if no Spotify presence was found
            if (spotifyUser.Presence == null || spotifyUser.Presence.Game == null || spotifyUser.Presence.Game.Name.ToLower() != "spotify")
            {
                await ctx.RespondAsync("Nothing is currently playing!");
                return;
            }

            string trackName = spotifyUser.Presence.Game.Details;
            string artistName = spotifyUser.Presence.Game.State;

            if (trackName == null || artistName == null)
            {
                await ctx.RespondAsync("Error: Couldn't recognize the current song");
                return;
            }

            // If current track and artist are the same, use cached embed instead
            if (trackName == Utils.cachedTrack && artistName == Utils.cachedArtist && Utils.cachedTrackEmbed != null)
            {
                await ctx.RespondAsync(embed: Utils.cachedTrackEmbed);
                return;
            }

            // Save searched track and artist
            JObject responseTrackJson = await Utils.FindAndCacheCurrentTrack(trackName, artistName);

            DiscordEmbed embed = Utils.BuildCachedTrackEmbed(responseTrackJson);

            if (embed == null)
            {
                await ctx.RespondAsync("Error: Something went wrong when creating the track embed");
                return;
            }

            await ctx.RespondAsync(embed: embed);
        }

        [Command("hi")]
        [Description("Hello!")]
        public async Task Hi(CommandContext ctx)
        {
            await ctx.RespondAsync($"👋 Hi, {ctx.User.Mention}!");
            return;
        }

        [Group("register", CanInvokeWithoutSubcommand = false)]
        [Description("Use with a subcommand to register for a tournament.")]
        public class RegisterCommands
        {
            private bool hsRegistrationOpen = false;
            private bool ctrRegistrationOpen = false;

            [Group("overwatch", CanInvokeWithoutSubcommand = false)]
            [Description("Register a team or player for the Overwatch tournament.")]
            public class OverwatchRegister
            {
                private bool owRegistrationOpen = false;

                [Command("toggle")]
                [Hidden] // let's hide this from the eyes of curious users
                [RequirePermissions(Permissions.Administrator)] // and restrict this to users who have appropriate permissions
                public async Task ToggleOWRegistration(CommandContext ctx)
                {
                    string open = "";

                    if (owRegistrationOpen)
                    {
                        owRegistrationOpen = false;
                        open = "closed";
                    }
                    else
                    {
                        owRegistrationOpen = true;
                        open = "open";
                    }

                    await ctx.RespondAsync("Overwatch registration is now " + open);
                }

                [Command("team")]
                [Description("Register a team for the Overwatch tournament.")]
                public async Task RegisterOWTeam(CommandContext ctx, [RemainingText, Description("The name of your team.")] string teamName)
                {
                    await ctx.TriggerTypingAsync();

                    if (!owRegistrationOpen)
                    {
                        await ctx.RespondAsync("Registration for the Overwatch tournament is closed!");
                        return;
                    }

                    if (teamName == null)
                    {
                        string error = ctx.User.Mention + " Please select a team name (!register overwatch team examplename)";
                        await ctx.RespondAsync(error);
                        return;
                    }

                    if (Program.fileOperations.isProcessing || isGettingData)
                    {
                        string error = ctx.User.Mention + " I am processing another operation, please wait.";
                        await ctx.RespondAsync(error);
                        return;
                    }
                    else
                    {
                        string response = Program.fileOperations.AddTeam(ctx.User, teamName, "Overwatch");

                        await ctx.RespondAsync(response);
                        isGettingData = false;
                        return;
                    }
                }

                [Command("player")]
                [Description("Register a user to your overwatch team. (You must have an existing team)")]
                public async Task RegisterOWPlayer(CommandContext ctx, [RemainingText, Description("User to add to your team.")] DiscordUser userToAdd)
                {
                    await ctx.TriggerTypingAsync();

                    if (!owRegistrationOpen)
                    {
                        await ctx.RespondAsync("Registration for the Overwatch tournament is closed!");
                        return;
                    }

                    if (userToAdd == null)
                    {
                        string error = ctx.User.Mention + " Please specify the Discord user to add (!register overwatch player @Wattumartti#1764)";
                        await ctx.RespondAsync(error);
                        return;
                    }

                    if (Program.fileOperations.isProcessing || isGettingData)
                    {
                        string error = ctx.User.Mention + " I am processing another operation, please wait.";
                        await ctx.RespondAsync(error);
                        return;
                    }
                    else
                    {
                        string response = Program.fileOperations.AddPlayerToTeam(ctx.User, userToAdd, "Overwatch");

                        await ctx.RespondAsync(response);
                        return;
                    }
                }
            }

            [Command("togglehs")]
            [Hidden] // let's hide this from the eyes of curious users
            [RequirePermissions(Permissions.Administrator)] // and restrict this to users who have appropriate permissions
            public async Task ToggleHSRegistration(CommandContext ctx)
            {
                string open = "";

                if (hsRegistrationOpen)
                {
                    hsRegistrationOpen = false;
                    open = "closed";
                }
                else
                {
                    hsRegistrationOpen = true;
                    open = "open";
                }

                await ctx.RespondAsync("Hearthstone registration is now " + open);
            }

            [Command("hearthstone")]
            [Description("Register to the Hearthstone tournament.")]
            public async Task HearthstoneRegister(CommandContext ctx)
            {
                await ctx.TriggerTypingAsync();

                if (!hsRegistrationOpen)
                {
                    await ctx.RespondAsync("Registration for the Hearthstone tournament is closed!");
                    return;
                }

                if (Program.fileOperations.isProcessing || isGettingData)
                {
                    string error = ctx.User.Mention + " I am processing another operation, please wait.";
                    await ctx.RespondAsync(error);
                    return;
                }
                else
                {
                    string response = Program.fileOperations.AddSoloPlayer(ctx.User, "Hearthstone");

                    await ctx.RespondAsync(response);
                    isGettingData = false;
                    return;
                }
            }

            [Command("togglectr")]
            [Hidden] // let's hide this from the eyes of curious users
            [RequirePermissions(Permissions.Administrator)] // and restrict this to users who have appropriate permissions
            public async Task ToggleCTRRegistration(CommandContext ctx)
            {
                string open = "";

                if (ctrRegistrationOpen)
                {
                    ctrRegistrationOpen = false;
                    open = "closed";
                }
                else
                {
                    ctrRegistrationOpen = true;
                    open = "open";
                }

                await ctx.RespondAsync("Crash Team Racing registration is now " + open);
            }

            [Command("ctr")]
            [Description("Register to the Crash Team Racing tournament.")]
            public async Task CtrRegister(CommandContext ctx)
            {
                await ctx.TriggerTypingAsync();

                if (!ctrRegistrationOpen)
                {
                    await ctx.RespondAsync("Registration for the Crash Team Racing tournament is closed!");
                    return;
                }

                if (Program.fileOperations.isProcessing || isGettingData)
                {
                    string error = ctx.User.Mention + " I am processing another operation, please wait.";
                    await ctx.RespondAsync(error);
                    return;
                }
                else
                {
                    string response = Program.fileOperations.AddSoloPlayer(ctx.User, "CTR");

                    await ctx.RespondAsync(response);
                    isGettingData = false;
                    return;
                }
            }
        }

        [Command("hsplayers")]
        [Description("Lists all registered Hearthstone players.")]
        public async Task HsPlayers(CommandContext ctx)
        {
            await ctx.TriggerTypingAsync();

            if (Program.fileOperations.isProcessing || isGettingData)
            {
                string error = ctx.User.Mention + " I am processing another operation, please wait.";
                await ctx.RespondAsync(error);
                return;
            }
            else
            {
                DiscordEmbedBuilder embed = Program.fileOperations.GetSoloPlayers(ctx.User, "Hearthstone");

                if (embed.Description == null)
                {
                    await ctx.RespondAsync(embed.Title);
                    return;
                }
                else
                {
                    await ctx.RespondAsync(embed: embed);
                    return;
                }
            }
        }

        [Command("ctrplayers")]
        [Description("Lists all registered Crash Team Racing players.")]
        public async Task CtrPlayers(CommandContext ctx)
        {
            await ctx.TriggerTypingAsync();

            if (Program.fileOperations.isProcessing || isGettingData)
            {
                string error = ctx.User.Mention + " I am processing another operation, please wait.";
                await ctx.RespondAsync(error);
                return;
            }
            else
            {
                DiscordEmbedBuilder embed = Program.fileOperations.GetSoloPlayers(ctx.User, "CTR");

                if (embed.Description == null)
                {
                    await ctx.RespondAsync(embed.Title);
                    return;
                }
                else
                {
                    await ctx.RespondAsync(embed: embed);
                    return;
                }
            }
        }

        [Command("owteams")]
        [Description("Lists all registered Overwatch teams and their captains")]
        public async Task OwTeams(CommandContext ctx)
        {
            await ctx.TriggerTypingAsync();

            if (Program.fileOperations.isProcessing || isGettingData)
            {
                string error = ctx.User.Mention + " I am processing another operation, please wait.";
                await ctx.RespondAsync(error);
                return;
            }
            else
            {
                DiscordEmbedBuilder embed = Program.fileOperations.GetOwTeams(ctx.User);

                if (embed.Description == null)
                {
                    await ctx.RespondAsync(embed.Title);
                    return;
                }
                else
                {
                    await ctx.RespondAsync(embed: embed);
                    return;
                }
            }
        }

        [Command("resetlocks")]
        [Description("Resets locking booleans")]
        [Hidden] // let's hide this from the eyes of curious users
        [RequirePermissions(Permissions.Administrator)] // and restrict this to users who have appropriate permissions
        public async Task ResetLocks(CommandContext ctx)
        {
            await ctx.TriggerTypingAsync();

            Program.fileOperations.isProcessing = false;
            isGettingData = false;

            await ctx.RespondAsync(ctx.User.Mention + " Locking booleans reset");
            return;
        }

        [Command("gb")]
        [Description("Link to our Facebook page")]
        public async Task GB(CommandContext ctx)
        {
            if (Program.fileOperations.isProcessing || isGettingData)
            {
                string error = ctx.User.Mention + " I am processing another operation, please wait.";
                await ctx.RespondAsync(error);
                return;
            }

            string link = "https://www.facebook.com/gamingbarracks";

            await ctx.Member.SendMessageAsync(link);
            return;
        }

        [Command("setschedule")]
        [Description("Sets the schedule image link")]
        [Hidden] // let's hide this from the eyes of curious users
        [RequirePermissions(Permissions.Administrator)] // and restrict this to users who have appropriate permissions
        public async Task SetScheduleLink(CommandContext ctx, [RemainingText, Description("The schedule image link")] string linkString)
        {
            await ctx.TriggerTypingAsync();

            if (Program.fileOperations.isProcessing || isGettingData)
            {
                string error = ctx.User.Mention + " I am processing another operation, please wait.";
                await ctx.Member.SendMessageAsync(error);
                return;
            }

            scheduleLink = linkString;

            await ctx.Member.SendMessageAsync(ctx.User.Mention + " Schedule link set!");
            return;
        }

        [Command("schedule")]
        [Description("Event schedule")]
        public async Task ShowSchedule(CommandContext ctx)
        {
            await ctx.TriggerTypingAsync();

            if (Program.fileOperations.isProcessing || isGettingData)
            {
                string error = ctx.User.Mention + " I am processing another operation, please wait.";
                await ctx.RespondAsync(error);
                return;
            }

            DiscordEmbedBuilder embed = new DiscordEmbedBuilder()
                .WithTitle("Event schedule:")
                .WithImageUrl(scheduleLink);

            await ctx.RespondAsync(embed: embed);
            return;
        }

        [Command("owstats")]
        [Description("Gets specified Battletag's Overwatch stats")]
        public async Task OverBuff(CommandContext ctx, [RemainingText] string battleTag)
        {
            await ctx.TriggerTypingAsync();

            if (string.IsNullOrEmpty(battleTag))
            {
                await ctx.RespondAsync(ctx.User.Mention + " Please use a correct battle tag. (!owstats Grimdonk#2975)");
                return;
            }

            string[] split = battleTag.Split('#');
            if (split.Length < 2 || split.Length > 2)
            {
                await ctx.RespondAsync(ctx.User.Mention + " Please use a correct battle tag. (!owstats Grimdonk#2975)");
                return;
            }

            if (Program.fileOperations.isProcessing || isGettingData)
            {
                string error = ctx.User.Mention + " I am processing another operation, please wait.";
                await ctx.RespondAsync(error);
                return;
            }

            DiscordEmbedBuilder embed = null;

            try
            {
                isGettingData = true;

                string tag = split[0] + "-" + split[1];

                RestClient restClient = new RestClient("https://ow-api.com/v1/stats/pc/eu/" + tag + "/profile");
                RestRequest request = new RestRequest(Method.GET);

                IRestResponse response = await restClient.ExecuteTaskAsync(request);
                string responseContent = response.Content;
                JObject json = JObject.Parse(responseContent);

                if (response.StatusCode != HttpStatusCode.OK)
                {
                    string error = ctx.User.Mention + " Error: Failed retrieving profile data.";
                    isGettingData = false;
                    await ctx.RespondAsync(error);
                    return;
                }

                JToken token = json["error"];

                if (token != null)
                {
                    string errorString = json.GetValue("error").ToString();

                    if (errorString != null)
                    {
                        string error = ctx.User.Mention + " Error: " + errorString;
                        isGettingData = false;
                        await ctx.RespondAsync(error);
                        return;
                    }
                }

                string iconString = json.GetValue("icon").ToString();
                string name = json.GetValue("name").ToString();
                string competitiveStatString = json.GetValue("competitiveStats").ToString();
                JObject competitiveStats = JObject.Parse(competitiveStatString);
                JObject competitiveGames = JObject.Parse(competitiveStats.GetValue("games").ToString());
                string competitiveWins = competitiveGames.GetValue("won").ToString();
                string competitivePlayed = competitiveGames.GetValue("played").ToString();
                string rating = json.GetValue("rating").ToString();
                string rankIconString = json.GetValue("ratingIcon").ToString();

                embed = new DiscordEmbedBuilder()
                        .WithTitle(split[0] + "#" + split[1])
                        .WithThumbnailUrl(iconString)
                        .AddField("Competitive games: ", competitivePlayed, true)
                        .AddField("Competitive wins: ", competitiveWins, true)
                        .WithTimestamp(DateTimeOffset.Now);

                if (!string.IsNullOrEmpty(rating) && !string.IsNullOrEmpty(rankIconString))
                {
                    embed.WithAuthor(name + "'s Overwatch profile", icon_url: rankIconString)
                        .AddField("Rating: ", rating);
                }
                else
                {
                    embed.WithAuthor(name + "'s Overwatch profile")
                        .AddField("Rating: ", "Not found");
                }
            }
            catch (Exception ex)
            {
                embed = new DiscordEmbedBuilder().WithTitle(ex.ToString());
            }

            isGettingData = false;
            await ctx.RespondAsync(embed: embed);
            return;
        }
    }
}
