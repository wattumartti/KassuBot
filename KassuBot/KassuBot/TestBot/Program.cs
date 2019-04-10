using System;
using DSharpPlus;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSharpPlus.CommandsNext;
using DSharpPlus.CommandsNext.Attributes;

namespace KassuBot
{
    class Program
    {
        static DiscordClient discord;
        static CommandsNextModule commands;
        public static FileOperations fileOperations;

        static void Main(string[] args)
        {
            MainAsync(args).ConfigureAwait(false).GetAwaiter().GetResult();
        }

        static async Task MainAsync(string[] args)
        {
            discord = new DiscordClient(new DiscordConfiguration
            {
                Token = "TOKEN",
                TokenType = TokenType.Bot,
                UseInternalLogHandler = true,
                LogLevel = LogLevel.Debug
            });

            commands = discord.UseCommandsNext(new CommandsNextConfiguration
            {
                StringPrefix = "!",
                EnableMentionPrefix = true,
                EnableDms = true        
            });

            commands.RegisterCommands<Commands>();

            fileOperations = new FileOperations();
            fileOperations.InitOverwatchSheet();
            fileOperations.InitHearthstoneSheet();
            fileOperations.InitCtrSheet();

            await discord.ConnectAsync();
            await Task.Delay(-1);
        }
    }
}
