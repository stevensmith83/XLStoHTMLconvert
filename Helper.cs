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
        public static List<string> headerList;

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
            headerList = new List<String>();
            headerList = ConfigurationManager.AppSettings["Header"].Split(';').ToList();
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
            string connectionString = "Provider=" + GetProvider() + ";Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO;';";

            OleDbConnection connection = new OleDbConnection(connectionString);
            OleDbCommand command = new OleDbCommand("Select * From [" + sheetName + "$]", connection);
            connection.Open();

            OleDbDataAdapter adapter = new OleDbDataAdapter(command);
            DataTable data = new DataTable();
            adapter.Fill(data);
            return data;
        }

        private static string GetProvider()
        {
            var reader = OleDbEnumerator.GetRootEnumerator();
            var provider = string.Empty;

            while (reader.Read())
            {
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    if (reader.GetName(i) == "SOURCES_NAME" && reader.GetValue(0).ToString().Contains("Microsoft.ACE.OLEDB"))
                    {
                        provider = reader.GetValue(i).ToString();
                        break;
                    }
                }                
            }

            reader.Close();
            return provider;
        }

        public static string Capitalize(string input)
        {
            if (!string.IsNullOrEmpty(input))
            {
                return char.ToUpper(input[0]) + input.Substring(1).ToLower();
            }

            return input;            
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
            return string.IsNullOrEmpty(data) ? data : data.Substring(0, data.IndexOf(";")) + Environment.NewLine + Capitalize(data.Substring(data.IndexOf(";") + 2));
        }

        internal static StringBuilder ConvertTableToHTML(DataGridView dataGridView)
        {
            StringBuilder html = new StringBuilder();

            html.AppendLine("<div><table align='center' border='0' cellpadding='2' cellspacing='2'>");
            html.AppendLine("<tbody><tr>");

            for (int col = 0; col < dataGridView.Columns.Count; col++)
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

                html.AppendLine("</strong></span></div><div style='text-align: center;'><span style='color: rgb(255, 255, 255);'><strong>" + prices[col - 1] + " Ft/adag</strong></span></div></td>");
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
                        if (firstCell)
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

        public static StringBuilder ConvertTableToBootstrapTable(DataGridView dataGridView)
        {
            StringBuilder html = new StringBuilder();

            html.AppendLine("<div class='table-responsive'>");
            html.AppendLine("\t<table class='table table-hover'>");
            html.Append(CreateHeader(dataGridView[0, 0].Value.ToString()));
            html.Append(CreateRows(dataGridView));
            html.AppendLine("\t</table>");
            html.AppendLine("</div>");
            return html;
        }

        private static string CreateHeader(string firstCell)
        {
            StringBuilder header = new StringBuilder();
            header.AppendLine("\t\t<thead>");
            header.AppendLine("\t\t\t<tr>");
            header.Append("\t\t\t\t<th style='text-align: center; vertical-align: middle;'>").Append(firstCell).AppendLine("</th>");

            foreach (string item in headerList)
            {
                header.Append("\t\t\t\t<th style='text-align: center; vertical-align: middle;'>").Append(item).AppendLine("</th>");
            }

            header.AppendLine("\t\t\t</tr>");
            header.AppendLine("\t\t</thead>");
            return header.ToString();
        }

        private static string CreateRows(DataGridView dataGridView)
        {
            StringBuilder rows = new StringBuilder();
            rows.AppendLine("\t\t<tbody>");            

            for (int row = 2; row < dataGridView.Rows.Count; row++)
            {
                rows.AppendLine("\t\t\t<tr>");
                Boolean firstCell = true;

                foreach (DataGridViewCell cell in dataGridView.Rows[row].Cells)
                {
                    string data = cell.Value.ToString();
                    data = data.Replace(System.Environment.NewLine, "<br><em>");

                    if (firstCell)
                    {
                        rows.Append("\t\t\t\t<th style='text-align: center; vertical-align: middle;'>").Append(data).AppendLine("</em></th>");
                        firstCell = false;
                    }
                    else
                    {
                        rows.Append("\t\t\t\t<td style='text-align: center; vertical-align: middle;'>").Append(data).AppendLine("</em></td>");
                    }
                }

                rows.AppendLine("\t\t\t</tr>");
            }
            
            rows.AppendLine("\t\t</tbody>");
            return rows.ToString();
        }
    }
}
