using System;
using System.Configuration;
using System.Text;
using System.Globalization;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace XLStoHTMLconvert
{
    public static class Helper
    {
        public static Dictionary<string, string> correction;
        public static List<string> prices;

        public static void InitDictionaryAndList()
        {
            correction = new Dictionary<string, string>();
            string input = ConfigurationManager.AppSettings["Correction"];

            if (input.Trim().Length > 0)
            {
                correction = input.TrimEnd(';').Split(';').ToDictionary(item => item.Split('=')[0], item => item.Split('=')[1]);
            }

            prices = new List<String>();
            prices = ConfigurationManager.AppSettings["Prices"].Split(';').ToList();
        }

        public static void SaveToDictionary(string key, string value)
        {
            correction.Add(key, value);
            SaveDictionaryToConfig();
        }

        public static void SaveDictionaryToConfig()
        {
            string configString = string.Join(";", correction.Select(x => x.Key + "=" + x.Value));
            Configuration config = ConfigurationManager.OpenExeConfiguration(System.Windows.Forms.Application.ExecutablePath);
            config.AppSettings.Settings.Remove("Correction");
            config.AppSettings.Settings.Add("Correction", configString);
            config.Save(ConfigurationSaveMode.Minimal);
        }

        public static DateTime FirstDateOfWeek(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

            DateTime firstThursday = jan1.AddDays(daysOffset);
            var cal = CultureInfo.CurrentCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            var weekNum = weekOfYear;

            if (firstWeek <= 1)
            {
                weekNum -= 1;
            }

            var result = firstThursday.AddDays(weekNum * 7);
            return result.AddDays(-3);
        }

        public static DataTable ImportXLS(string fileName)
        {
            string sheetName = ConfigurationManager.AppSettings["SheetName"];
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO;';";

            OleDbConnection connection = new OleDbConnection(connectionString);
            OleDbCommand command = new OleDbCommand("Select * From [" + sheetName + "$]", connection);
            connection.Open();

            OleDbDataAdapter adapter = new OleDbDataAdapter(command);
            DataTable data = new DataTable();
            adapter.Fill(data);
            return data;
        }

        public static string Capitalize(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                throw new ArgumentException("ARGH!");
            }

            return char.ToUpper(input[0]) + input.Substring(1).ToLower();
        }

        public static string FormatFirstCell(DateTime day)
        {
            StringBuilder builder = new StringBuilder();

            builder.Append(day.ToString("dddd", new CultureInfo("hu-HU")).ToUpper()).Append(Environment.NewLine);
            builder.Append(day.ToString("MMMM", new CultureInfo("hu-HU"))).Append(" ");
            builder.Append(day.ToString("dd", new CultureInfo("hu-HU"))).Append(".");

            return builder.ToString();
        }

        public static string FormatCellData(string data)
        {
            data = Capitalize(data.Trim());
            data = Regex.Replace(data, "[ ]{2,}", " ");
            data = Regex.Replace(data, @"\,(?! |$)", ", ");
            data = correction.ContainsKey(data) ? correction[data] : data;
            return data;
        }

        public static string Format7thCell(string data, string firstCell, string secondCell)
        {
            data = data.Replace("L1,", firstCell + ";").Replace("L2,", secondCell + ";");
            return data.Substring(0, data.IndexOf(";")) + Environment.NewLine + Capitalize(data.Substring(data.IndexOf(";") + 2));
        }

        internal static StringBuilder ConvertTableToHTML(DataGridView dataGridView)
        {
            StringBuilder html = new StringBuilder();
            
            html.AppendLine("<div><table align='center' border='0' cellpadding='2' cellspacing='2'>");
            html.AppendLine("<tbody><tr>");

            for(int col = 0; col < dataGridView.Columns.Count; col++)
            {
                if (!string.IsNullOrEmpty(dataGridView[col, 0].Value.ToString()))
                {
                    html.Append("<td ");
                    switch (col)
                    {
                        case 0:
                            html.Append("rowspan='2' ");
                            break;
                        case 1:
                        case 5:
                            html.Append("colspan='2' ");
                            break;
                    }

                    html.AppendLine("style='text-align: center; background-color: rgb(153, 0, 0);'><div><span style = 'color: rgb(255, 255, 255);'><strong>" + dataGridView[col, 0].Value + "</strong></span></div></td>");
                }                
            }

            html.AppendLine("</tr><tr>");

            for (int col = 1; col < dataGridView.Columns.Count; col++)
            {
                html.Append("<td style='background-color: rgb(153, 0, 0);'><div style='text-align: center;'><span style='color: rgb(255, 255, 255);'><strong>");
                switch (col)
                {
                    case 1:
                    case 5:
                        html.Append("I.");
                        break;
                    case 2:
                    case 6:
                        html.Append("II.");
                        break;
                }

                html.AppendLine("</strong></span></div><div style='text-align: center;'><span style='color: rgb(255, 255, 255);'><strong>" + prices[col-1] + " Ft/adag</strong></span></div></td>");
            }

            html.AppendLine("</tr>");
            
            for (int row = 2; row < dataGridView.Rows.Count; row++)
            {
                html.AppendLine("<tr>");
                Boolean firstCell = true;

                foreach (DataGridViewCell cell in dataGridView.Rows[row].Cells)
                {
                    string color = firstCell ? "rgb(153, 0, 0)" : "rgb(255, 102, 0)";
                    string data = cell.Value.ToString();
                    string emTag = "";

                    if (data.Contains(System.Environment.NewLine))
                    {
                        if(firstCell)
                        {
                            data = "<strong>" + data.Replace(System.Environment.NewLine, "</strong></span></div><div><span style='color: rgb(255, 255, 255);'>");
                        }
                        else
                        {
                            data = data.Replace(System.Environment.NewLine, "</span></em></div><div><em><span style='color:#ffffff;'>");
                            emTag = "</em>";
                        }
                    }

                    html.AppendLine("<td style='text-align: center; background-color: " + color + ";'><div><span style='color:#ffffff;'>" + data + "</span>" + emTag + "</div></td>");
                    firstCell = false;
                }

                html.AppendLine("</tr>");
            }
            
            html.AppendLine("</tbody></table></div>");
            return html;
        }

        public static string Format8thCell(string data)
        {
            return data.Replace("e:", Environment.NewLine + "E: ").Replace("szh:", Environment.NewLine + "Szh: "); ;
        }
    }
}
