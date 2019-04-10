using DSharpPlus.Entities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace KassuBot
{
    class FileOperations
    {
        public bool isProcessing = false;
        private string teamListFilePath = @"FILEPATH";
        public int teamCountOverwatch = 0;
        public int hearthstonePlayerCount = 0;
        public int ctrPlayerCount = 0;
        private int startRow = 3;

        public void InitOverwatchSheet()
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Workbooks workBooks = null;
            Excel._Worksheet oSheet = null;
            Excel.Sheets sheets = null;

            try
            {
                isProcessing = true;

                oXL = new Excel.Application();
                workBooks = oXL.Workbooks;
                oWB = workBooks.Open(teamListFilePath);
                sheets = oWB.Worksheets;
                oSheet = (Excel._Worksheet)sheets["Overwatch"];

                for (int i = 0; i < 100; i++)
                {
                    if (string.Equals(oSheet.Cells[startRow + (i * 2), 1].Value, null))
                    {
                        teamCountOverwatch = i + 1;
                        break;
                    }
                }
            }
            finally
            {
                if (oWB != null)
                {
                    foreach (Excel.Workbook _workbook in oXL.Workbooks)
                    {
                        _workbook.Close();
                    }

                    oXL.Quit();
                    oXL = null;
                    var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (var p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }
            }
            isProcessing = false;
        }

        public void InitHearthstoneSheet()
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Workbooks workBooks = null;
            Excel._Worksheet oSheet = null;
            Excel.Sheets sheets = null;

            try
            {
                isProcessing = true;

                oXL = new Excel.Application();
                workBooks = oXL.Workbooks;
                oWB = workBooks.Open(teamListFilePath);
                sheets = oWB.Worksheets;
                oSheet = (Excel._Worksheet)sheets["Hearthstone"];

                for (int i = 0; i < 100; i++)
                {
                    if (string.Equals(oSheet.Cells[startRow + (i * 2), 1].Value, null))
                    {
                        hearthstonePlayerCount = i + 1;
                        break;
                    }
                }
            }
            finally
            {
                if (oWB != null)
                {
                    foreach (Excel.Workbook _workbook in oXL.Workbooks)
                    {
                        _workbook.Close();
                    }

                    oXL.Quit();
                    oXL = null;
                    var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (var p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }
            }
            isProcessing = false;
        }

        public void InitCtrSheet()
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Workbooks workBooks = null;
            Excel._Worksheet oSheet = null;
            Excel.Sheets sheets = null;

            try
            {
                isProcessing = true;

                oXL = new Excel.Application();
                workBooks = oXL.Workbooks;
                oWB = workBooks.Open(teamListFilePath);
                sheets = oWB.Worksheets;
                oSheet = (Excel._Worksheet)sheets["CTR"];

                for (int i = 0; i < 100; i++)
                {
                    if (string.Equals(oSheet.Cells[startRow + (i * 2), 1].Value, null))
                    {
                        ctrPlayerCount = i + 1;
                        break;
                    }
                }
            }
            finally
            {
                if (oWB != null)
                {
                    foreach (Excel.Workbook _workbook in oXL.Workbooks)
                    {
                        _workbook.Close();
                    }

                    oXL.Quit();
                    oXL = null;
                    var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (var p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }
            }
            isProcessing = false;
        }

        public string AddTeam(DiscordUser user, string teamName, string sheetName)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Workbooks workBooks = null;
            Excel._Worksheet oSheet = null;
            Excel.Sheets sheets = null;
            Excel.Range oRng = null;

            string returnString = null;

            try
            {
                isProcessing = true;

                oXL = new Excel.Application();
                workBooks = oXL.Workbooks;
                oWB = workBooks.Open(teamListFilePath);
                sheets = oWB.Worksheets;
                oSheet = (Excel._Worksheet)sheets[sheetName];

                bool emptyCellFound = false;
                bool teamNameFound = false;
                bool hasTeam = false;
                int count = 0;
                string ownerTag = user.Username + " " + user.Mention;
                while (!emptyCellFound && !teamNameFound && !hasTeam)
                {
                    if (string.Equals(oSheet.Cells[startRow + (count * 2), 1].Value, null))
                    {
                        emptyCellFound = true;
                    }
                    else if (string.Equals(oSheet.Cells[startRow + (count * 2), 1].Value, teamName))
                    {
                        teamNameFound = true;
                    }
                    else if (string.Equals(oSheet.Cells[startRow + (count * 2), 2].Value, ownerTag))
                    {
                        hasTeam = true;
                    }
                    else
                    {
                        count++;
                    }
                }

                if (teamNameFound)
                {
                    returnString = teamName + " has already been registered!";
                }
                else if (hasTeam)
                {
                    returnString = user.Mention + " You have already registered a team!";
                }
                else if (emptyCellFound)
                {
                    oSheet.Cells[startRow + (count * 2), 1].Value = teamName;
                    oRng = oSheet.get_Range("A2", "N2");
                    oRng.EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    oRng.EntireColumn.Font.Bold = true;
                    oRng.EntireColumn.Font.Size = 15;
                    oRng.EntireColumn.AutoFit();

                    oRng = oSheet.Rows[startRow + (count * 2)];
                    Excel.Borders borders = oRng.Borders;
                    borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 4d;

                    oRng = oSheet.Cells[startRow + (count * 2), 1];
                    oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 3d;

                    oSheet.Cells[startRow + (count * 2), 2].Value = ownerTag;

                    returnString = user.Mention + " " + teamName + " registered!";

                    oWB.Save();
                }

            }
            catch (Exception ex)
            {
                string exceptionString = ex.ToString();
                if (exceptionString.Contains("Microsoft.Office.Interop.Excel.Workbooks.Open"))
                {
                    return user.Mention + " I am processing another operation, please wait.";
                }
                return "Error: Something went wrong, please contact an admin.";
            }
            finally
            {
                if (oWB != null)
                {
                    foreach (Excel.Workbook _workbook in oXL.Workbooks)
                    {
                        _workbook.Close();
                    }

                    oXL.Quit();
                    oXL = null;
                    var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (var p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }
            }

            switch (sheetName)
            {
                case "Overwatch":
                    teamCountOverwatch++;
                    break;
            }

            isProcessing = false;

            return returnString;
        }

        public string AddPlayerToTeam(DiscordUser user, DiscordUser userToAdd, string sheetName)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Workbooks workBooks = null;
            Excel._Worksheet oSheet = null;
            Excel.Sheets sheets = null;
            Excel.Range oRng = null;

            string returnString = null;

            try
            {
                isProcessing = true;

                oXL = new Excel.Application();
                workBooks = oXL.Workbooks;
                oWB = workBooks.Open(teamListFilePath);
                sheets = oWB.Worksheets;
                oSheet = (Excel._Worksheet)sheets[sheetName];

                bool emptyCellFound = false;
                bool teamNameFound = false;
                bool userFound = false;
                int teamRow = 0;
                string teamName = null;
                string ownerTag = user.Username + " " + user.Mention;
                string addedUserTag = userToAdd.Username + " " + userToAdd.Mention;

                for (int i = 0; i < teamCountOverwatch; i++)
                {
                    if (ownerTag.Equals(oSheet.Cells[startRow + (i * 2), 2].Value))
                    {
                        teamNameFound = true;
                        teamRow = startRow + (i * 2);
                        teamName = oSheet.Cells[startRow + (i * 2), 1].Value;
                        oRng = oSheet.Rows[teamRow];
                        break;
                    }            
                    
                }
                if (!teamNameFound || teamName == null)
                {
                    returnString = "Your team was not found!";
                }
                else if (teamNameFound)
                {
                    int foundColumn = 0;      
                    for (int i = 1; i < 100; i++)
                    {
                        if (string.Equals(oSheet.Cells[teamRow, i].Value, null))
                        {
                            emptyCellFound = true;
                            foundColumn = i;
                            break;
                        }
                        else if (string.Equals(oSheet.Cells[teamRow, i].Value, addedUserTag))
                        {
                            userFound = true;
                            break;
                        }
                    }
                    if (userFound)
                    {
                        returnString = userToAdd.Mention + " is already in this team!";
                    }
                    else if (emptyCellFound)
                    {
                        oSheet.Cells[teamRow, foundColumn].Value = addedUserTag;
                        oWB.Save();
                        returnString = userToAdd.Mention + " added to team " + teamName;
                    }
                }

            }
            catch (Exception ex)
            {
                string exceptionString = ex.ToString();
                if (exceptionString.Contains("Microsoft.Office.Interop.Excel.Workbooks.Open"))
                {
                    return user.Mention + " I am processing another operation, please wait.";
                }
                return "Error: Something went wrong, please contact an admin.";
            }
            finally
            {
                if (oWB != null)
                {
                    foreach (Excel.Workbook _workbook in oXL.Workbooks)
                    {
                        _workbook.Close();
                    }

                    oXL.Quit();
                    oXL = null;
                    var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (var p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }
            }

            isProcessing = false;
            return returnString;
        }

        public string ChangeTeamName(DiscordUser user, string teamName, string sheetName, out string oldName)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Workbooks workBooks = null;
            Excel._Worksheet oSheet = null;
            Excel.Sheets sheets = null;

            bool teamNameFound = false;
            int teamRow = 0;
            string returnString = null;
            string ownerTag = user.Username + " " + user.Mention;
            string oldTeamName = null;

            try
            {
                isProcessing = true;

                oXL = new Excel.Application();
                workBooks = oXL.Workbooks;
                oWB = workBooks.Open(teamListFilePath);
                sheets = oWB.Worksheets;
                oSheet = (Excel._Worksheet)sheets[sheetName];

                for (int i = 0; i < teamCountOverwatch; i++)
                {
                    if (ownerTag.Equals(oSheet.Cells[startRow + (i * 2), 2].Value))
                    {
                        teamNameFound = true;
                        teamRow = startRow + (i * 2);
                        break;
                    }
                }

                if (!teamNameFound)
                {
                    returnString = user.Mention + " You do not own a team!";
                }
                else
                {
                    oldTeamName = oSheet.Cells[teamRow, 1].Value;
                    oSheet.Cells[teamRow, 1].Value = teamName;
                    oWB.Save();

                    returnString = user.Mention + " Team's name changed from " + '"' + oldTeamName + '"' + " to " + '"' + teamName + '"';
                }
            }
            catch (Exception ex)
            {
                string exceptionString = ex.ToString();
                if (exceptionString.Contains("Microsoft.Office.Interop.Excel.Workbooks.Open"))
                {
                    oldName = null;
                    return user.Mention + " I am processing another operation, please wait.";
                }
                oldName = null;
                return "Error: Something went wrong, please contact an admin.";
            }
            finally
            {
                if (oWB != null)
                {
                    foreach (Excel.Workbook _workbook in oXL.Workbooks)
                    {
                        _workbook.Close();
                    }

                    oXL.Quit();
                    oXL = null;
                    var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (var p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }
            }

            isProcessing = false;
            oldName = oldTeamName;
            return returnString;
        }

        public DiscordEmbedBuilder GetOwTeams(DiscordUser user)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Workbooks workBooks = null;
            Excel._Worksheet oSheet = null;
            Excel.Sheets sheets = null;

            string teamNames = null;
            string captainNames = null;
            bool hasTeams = false;

            DiscordEmbedBuilder embed = null;

            try
            {
                isProcessing = true;

                oXL = new Excel.Application();
                workBooks = oXL.Workbooks;
                oWB = workBooks.Open(teamListFilePath);
                sheets = oWB.Worksheets;
                oSheet = (Excel._Worksheet)sheets["Overwatch"];

                for (int i = 0; i < teamCountOverwatch; i++)
                {
                    if (!string.Equals(oSheet.Cells[startRow + (i * 2), 1].Value, null))
                    {
                        hasTeams = true;
                        teamNames += oSheet.Cells[startRow + (i * 2), 1].Value + "\n\n";

                        string cellValue = oSheet.Cells[startRow + (i * 2), 2].Value;
                        string[] splitMention = cellValue.Split(' ');
                        captainNames += splitMention[1] + "\n\n";
                    }
                }

                if (!hasTeams)
                {
                    embed = new DiscordEmbedBuilder()
                        .WithTitle("Registered Overwatch Teams:")
                        .WithDescription("A list of all registered Overwatch teams and their captains.")
                        .WithColor(new DiscordColor(0x106FD4))
                        .WithTimestamp(DateTimeOffset.Now)
                        .AddField("Error", "There are currently no registered teams.");
                }
                else
                {
                    embed = new DiscordEmbedBuilder()
                        .WithTitle("Registered Overwatch Teams:")
                        .WithDescription("A list of all registered Overwatch teams and their captains.")
                        .WithColor(new DiscordColor(0x106FD4))
                        .WithTimestamp(DateTimeOffset.Now)
                        .AddField("Team Name", teamNames, true)
                        .AddField("Captain", captainNames, true);
                }
            }
            catch (Exception ex)
            {
                string exceptionString = ex.ToString();
                string error = null;
                if (exceptionString.Contains("Microsoft.Office.Interop.Excel.Workbooks.Open"))
                {
                    error = user.Mention + " I am processing another operation, please wait.";
                }
                else
                {
                    error = "Error: Something went wrong, please contact an admin.";
                }

                if (error != null)
                {
                    embed = new DiscordEmbedBuilder().WithTitle(error);
                }

            }
            finally
            {
                if (oWB != null)
                {
                    foreach (Excel.Workbook _workbook in oXL.Workbooks)
                    {
                        _workbook.Close();
                    }

                    oXL.Quit();
                    oXL = null;
                    var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (var p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }
            }

            isProcessing = false;
            return embed;
        }

        public string AddSoloPlayer(DiscordUser user, string sheetName)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Workbooks workBooks = null;
            Excel._Worksheet oSheet = null;
            Excel.Sheets sheets = null;
            Excel.Range oRng = null;

            string returnString = null;

            try
            {
                isProcessing = true;

                oXL = new Excel.Application();
                workBooks = oXL.Workbooks;
                oWB = workBooks.Open(teamListFilePath);
                sheets = oWB.Worksheets;
                oSheet = (Excel._Worksheet)sheets[sheetName];

                bool emptyCellFound = false;
                bool nameFound = false;
                bool hasPlayer = false;
                int count = 0;
                string userTag = user.Username + " " + user.Mention;
                while (!emptyCellFound && !nameFound && !hasPlayer)
                {
                    if (string.Equals(oSheet.Cells[startRow + (count * 2), 1].Value, null))
                    {
                        emptyCellFound = true;
                    }
                    else if (string.Equals(oSheet.Cells[startRow + (count * 2), 1].Value, userTag))
                    {
                        hasPlayer = true;
                    }
                    else
                    {
                        count++;
                    }
                }

                if (hasPlayer)
                {
                    returnString = user.Mention + " You have already registered to the tournament!";
                }
                else if (emptyCellFound)
                {
                    oSheet.Cells[startRow + (count * 2), 1].Value = userTag;
                    oRng = oSheet.get_Range("A2", "N2");
                    oRng.EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    oRng.EntireColumn.Font.Bold = true;
                    oRng.EntireColumn.Font.Size = 15;
                    oRng.EntireColumn.AutoFit();

                    oRng = oSheet.Rows[startRow + (count * 2)];
                    Excel.Borders borders = oRng.Borders;
                    borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 4d;

                    oRng = oSheet.Cells[startRow + (count * 2), 1];
                    oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 3d;

                    returnString = user.Mention + " registered!";

                    oWB.Save();
                }

            }
            catch (Exception ex)
            {
                string exceptionString = ex.ToString();
                if (exceptionString.Contains("Microsoft.Office.Interop.Excel.Workbooks.Open"))
                {
                    return user.Mention + " I am processing another operation, please wait.";
                }
                return "Error: Something went wrong, please contact an admin.";
            }
            finally
            {
                if (oWB != null)
                {
                    foreach (Excel.Workbook _workbook in oXL.Workbooks)
                    {
                        _workbook.Close();
                    }

                    oXL.Quit();
                    oXL = null;
                    var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (var p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }
            }

            switch (sheetName)
            {
                case "Hearthstone":
                    hearthstonePlayerCount++;
                    break;
                case "CTR":
                    ctrPlayerCount++;
                    break;
            }
            hearthstonePlayerCount++;

            isProcessing = false;

            return returnString;
        }

        public DiscordEmbedBuilder GetSoloPlayers(DiscordUser user, string sheetName)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel.Workbooks workBooks = null;
            Excel._Worksheet oSheet = null;
            Excel.Sheets sheets = null;

            string userNames = null;
            bool hasUsers = false;
            DiscordEmbedBuilder embed = null;

            int playerCount = 0;

            switch (sheetName)
            {
                case "Hearthstone":
                    playerCount = hearthstonePlayerCount;
                    break;
                case "CTR":
                    playerCount = ctrPlayerCount;
                    break;
            }

            try
            {
                isProcessing = true;

                oXL = new Excel.Application();
                workBooks = oXL.Workbooks;
                oWB = workBooks.Open(teamListFilePath);
                sheets = oWB.Worksheets;
                oSheet = (Excel._Worksheet)sheets[sheetName];

                for (int i = 0; i < playerCount; i++)
                {
                    if (!string.Equals(oSheet.Cells[startRow + (i * 2), 1].Value, null))
                    {
                        hasUsers = true;
                        string cellValue = oSheet.Cells[startRow + (i * 2), 1].Value;
                        string[] splitMention = cellValue.Split(' ');
                        userNames += splitMention[1] + "\n\n";
                    }
                }

                if (!hasUsers)
                {
                    embed = new DiscordEmbedBuilder()
                        .WithTitle("Registered " + sheetName + " Players:")
                        .WithDescription("A list of all registered " + sheetName + " players.")
                        .WithColor(new DiscordColor(0x106FD4))
                        .WithTimestamp(DateTimeOffset.Now)
                        .AddField("Error", "There are currently no registered players.");
                }
                else
                {
                    embed = new DiscordEmbedBuilder()
                        .WithTitle("Registered " + sheetName + " Players:")
                        .WithDescription("A list of all registered " + sheetName + " players.")
                        .WithColor(new DiscordColor(0x106FD4))
                        .WithTimestamp(DateTimeOffset.Now)
                        .AddField("Player Name", userNames);
                }
            }
            catch (Exception ex)
            {
                string exceptionString = ex.ToString();
                string error = null;
                if (exceptionString.Contains("Microsoft.Office.Interop.Excel.Workbooks.Open"))
                {
                    error = user.Mention + " I am processing another operation, please wait.";
                }
                else
                {
                    error = "Error: Something went wrong, please contact an admin.";
                }

                if (error != null)
                {
                    embed = new DiscordEmbedBuilder().WithTitle(error);
                }

            }
            finally
            {
                if (oWB != null)
                {
                    foreach (Excel.Workbook _workbook in oXL.Workbooks)
                    {
                        _workbook.Close();
                    }

                    oXL.Quit();
                    oXL = null;
                    var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (var p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }
            }

            isProcessing = false;
            return embed;
        }
    }
}
