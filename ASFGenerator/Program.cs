using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
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

        static int rowskip = 0;
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

            Console.Write("Enter game for idle (default 730):");
            string result = Console.ReadLine();
            if (!Int32.TryParse(result, out idleGame)) idleGame = 730;

            Console.Write("Enter row SKIP number (skip headers):");
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out rowskip)) rowskip = 0;

            Console.Write("Enter name column number:");
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out name)) name = 999999999;

            Console.Write("Enter login column number:");
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out login)) login = 2;

            Console.Write("Enter password column number:");
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out password)) password = 3;

            Console.Write("Enter token column number:");
            result = Console.ReadLine();
            if (!Int32.TryParse(result, out token)) token = 4;

            if (name == 999999999) name = login;

            List<Model> model = ReadFromExcel(steam);

            foreach (var account in model)
            {
                Console.WriteLine("{0} {1} {2} {3}", account.name, account.login, account.password, account.token);

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

                foreach (var file in Directory.GetFiles(mafiles))
                {
                    string fileText = File.ReadAllText(file);
                    if (fileText.Contains(String.Format("account_name\":\"{0}", account.login.ToLower())))
                    {
                        string tofile = String.Format("{0}\\{1}.maFile", config, account.name.Replace(" ", ""));
                        File.Copy(file, tofile, true);
                        Console.WriteLine("{0} SDA Cкопирован!", account.name);
                        break;
                    }
                }
            }

            Console.ReadLine();
        }

        static List<Model> ReadFromExcel(string FilePath)
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

                for (int r = rowskip; r < rowCount; r++) 
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
                        if (worksheet.Cells[r, token].Value == null) continue;
                        model.token = worksheet.Cells[r, token].Value.ToString();
                        list.Add(model);
                    }
                    catch { continue; }
                }

                return list;
            }
        }
    }
}
