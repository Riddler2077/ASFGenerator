using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Media3D;

namespace ASFGenerator
{
    internal class Program
    {
        static string steam = String.Format("{0}\\steam.xlsx", Environment.CurrentDirectory);

        static string mafiles = String.Format("{0}\\maFiles", Environment.CurrentDirectory);

        static string config = String.Format("{0}\\config", Environment.CurrentDirectory);

        static int rowskip = 1;
        static int name = 0;
        static int login = 0;
        static int password = 0;
        static int token = 0;
        static int idleGame = 730;

        static void Main(string[] args)
        {
            if (!Directory.Exists(mafiles)) Directory.CreateDirectory(mafiles);

            if (!Directory.Exists(config)) Directory.CreateDirectory(config);

            if (!File.Exists(steam)) { Console.WriteLine("steam.xlsx not found!"); Console.ReadLine(); return; }

            string result = string.Empty;

            Dictionary<string, string> language = new Dictionary<string, string>();

            Console.Write("Select language (default EN):");
            result = Console.ReadLine();
            if (result.ToLower().StartsWith("en") || string.IsNullOrEmpty(result) || string.IsNullOrWhiteSpace(result)) language = english;
            else if(result.ToLower().StartsWith("ru")) language = russian;
            else language = english;

            Console.Write(language["gameForIdle"]);
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out idleGame)) idleGame = 730;

            Console.Write(language["skipHeaders"]);
            Console.SetCursorPosition(language["skipHeaders"].Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)[0].Length, (Console.CursorTop - 1));
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out rowskip)) rowskip = 1;
            ClearNote(language["skipHeaders"]);

            Console.Write(language["columnName"]);
            Console.SetCursorPosition(language["columnName"].Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)[0].Length, (Console.CursorTop - 1));
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out name)) name = 999999999;
            ClearNote(language["columnName"]);

            Console.Write(language["columnLogin"]);
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out login)) login = 1;

            Console.Write(language["columnPassword"]);
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out password)) password = 2;

            Console.Write(language["columnToken"]);
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out token)) token = 3;

            if (name == 999999999) name = login;

            List<Model> model = ReadFromExcel(steam);

            foreach (var account in model)
            {
                if (language["language"] == "en") Console.WriteLine("Created {0}.json with data {1} {2} {3}", account.name, account.login, account.password, account.token);
                else Console.WriteLine("Создан {0}.json с данными {1} {2} {3}", account.name, account.login, account.password, account.token);

                string filename = String.Format("{0}\\{1}.json", config, account.name.Replace(" ", ""));
                using (var writer = new StreamWriter(filename, false))
                {
                    writer.WriteLine("{");
                    writer.WriteLine("  \"AcceptGifts\": true,");
                    writer.WriteLine("  \"AutoSteamSaleEvent\": true,");
                    writer.WriteLine("  \"Enabled\": true,");
                    writer.WriteLine("  \"GamesPlayedWhileIdle\": [");
                    writer.WriteLine("    " + idleGame.ToString());
                    writer.WriteLine("  ],");
                    writer.WriteLine("  \"RedeemingPreferences\": 7,");
                    writer.WriteLine("  \"SendOnFarmingFinished\": true,");
                    writer.WriteLine("  \"SteamLogin\": \"" + account.login + "\",");
                    writer.WriteLine("  \"SteamPassword\": \"" + account.password + "\",");
                    writer.WriteLine("  \"SteamTradeToken\": \"" + account.token + "\",");
                    writer.WriteLine("  \"TradingPreferences\": 1");
                    writer.WriteLine("}");
                }

                //Mb linq?
                foreach (var file in Directory.GetFiles(mafiles))
                {
                    string fileText = File.ReadAllText(file);
                    if (fileText.Contains(String.Format("account_name\":\"{0}\"", account.login.ToLower())))
                    {
                        string tofile = String.Format("{0}\\{1}.maFile", config, account.name.Replace(" ", ""));
                        File.Copy(file, tofile, true);
                        if (language["language"] == "en") Console.WriteLine("{0} maFile found. Copied to config.", account.name);
                        else Console.WriteLine("{0} maFile найден. Скопирован в config.", account.name);
                        break;
                    }
                }
            }
            
            Console.ReadLine();
        }

        private static void ClearNote(string note) 
        {
            Console.SetCursorPosition(note.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries)[1].Length, Console.CursorTop);
            do { Console.Write("\b \b"); } while (Console.CursorLeft > 0);
        }

        private static Dictionary<string, string> english = new Dictionary<string, string>
        {
            {"language", "en" },
            {"gameForIdle", "Enter the game number for idle (default 730):"},
            {"skipHeaders", String.Format("How many lines need to skip from the top (default 1):{0}Note: value to skip headers. If there are no headers - leave input is empty", Environment.NewLine) },
            {"columnName", String.Format("Enter the number of the column with nicknames:{0}Note: leave the input blank if you want the nickname to be the same as the login", Environment.NewLine) },
            {"columnLogin", "Enter the number of the column with logins (default 1):" },
            {"columnPassword", "Enter the number of the column with passwords (default 2):" },
            {"columnToken", "Enter the number of the column with tokens (default 3):" }
        };

        private static Dictionary<string, string> russian = new Dictionary<string, string>
        {
            {"language", "ru" },
            {"gameForIdle", "Введите номер игры для фарма часов (по умолчанию 730):"},
            {"skipHeaders", String.Format("Сколько строк нужно отступить сверху (по умолчанию 1):{0}Подсказка: значение отступа для пропуска заголовков. Если заголовков нет - оставьте строку пустой", Environment.NewLine) },
            {"columnName", String.Format("Введите номер столбца с никами:{0}Подсказка: оставьте ввод пустым если хотите чтобы ник был таким же как и логин", Environment.NewLine) },
            {"columnLogin", "Введите номер столбца с логинами (по умолчанию 1):" },
            {"columnPassword", "Введите номер столбца с паролями (по умолчанию 2):" },
            {"columnToken", "Введите номер столбца с токенами (по умолчанию 3):" }
        };

        #region ReadFromExcel
        private static List<Model> ReadFromExcel(string FilePath)
        {
            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count

                List<Model> list = new List<Model>();

                for (int r = rowskip; r < (rowCount + 1); r++) 
                {
                    Model model = new Model();
                    try
                    {
                        if (worksheet.Cells[r, name].Value == null) continue;
                        model.name = worksheet.Cells[r, name].Value.ToString();
                        if (worksheet.Cells[r, login].Value == null) continue;
                        model.login = worksheet.Cells[r, login].Value.ToString();
                        if (worksheet.Cells[r, password].Value == null) continue;
                        model.password = worksheet.Cells[r, password].Value.ToString();
                        if (worksheet.Cells[r, token].Value == null) { list.Add(model); continue; }
                        model.token = worksheet.Cells[r, token].Value.ToString();
                        list.Add(model);
                    }
                    catch { continue; }
                }

                return list;
            }
        }

        #endregion
    }
}
