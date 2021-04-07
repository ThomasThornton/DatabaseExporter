using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Xml.Linq;

namespace Database_Exporter
{
    public partial class MainWindow : Window
    {
        string db_filename, db_type, db_table;
        List<Button> buttons_list = new List<Button>();
        List<string> table_list = new List<string>();
        List<string> second_list = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
        }

        // Eksporter til excel fil
        private void ExportToExcel(List<string> list)
        {
            ExcelPackage excelpackage = new ExcelPackage();
            excelpackage.Workbook.Properties.Created = DateTime.Now;
            string type = db_type.Split(',')[0];

            if (type.Equals("History"))
            {
                if (db_table.Equals("Webhistorik"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Webhistorik seneste");

                    worksheet.Cells["A1"].Value = "URL";
                    worksheet.Cells["B1"].Value = "Titel";
                    worksheet.Cells["C1"].Value = "Senest besøgt (UTC)";
                    worksheet.Cells["A1:C1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(1).Width = 100;
                    worksheet.Column(1).Style.WrapText = true;
                    worksheet.Column(2).Width = 100;
                    worksheet.Column(2).Style.WrapText = true;
                    worksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                    if (second_list.Count != 0)
                    {
                        ExcelWorksheet worksheet2 = excelpackage.Workbook.Worksheets.Add("Webhistorik alle");

                        worksheet2.Cells["A1"].Value = "URL";
                        worksheet2.Cells["B1"].Value = "Titel";
                        worksheet2.Cells["C1"].Value = "Senest besøgt (UTC)";
                        worksheet2.Cells["A1:C1"].Style.Font.Bold = true;

                        for (int i = 0; i < second_list.Count; i++)
                        {
                            string[] columns = second_list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                            worksheet2.Cells[i + 2, 1].Value = columns[0];
                            worksheet2.Cells[i + 2, 2].Value = columns[1];
                            worksheet2.Cells[i + 2, 3].Value = columns[2];
                        }

                        worksheet2.Cells[worksheet.Dimension.Address].AutoFitColumns();
                        worksheet2.Column(1).Width = 100;
                        worksheet2.Column(1).Style.WrapText = true;
                        worksheet2.Column(2).Width = 100;
                        worksheet2.Column(2).Style.WrapText = true;
                        worksheet2.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                        second_list.Clear();
                    }
                }

                if (db_table.Equals("Downloads"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Downloads");

                    worksheet.Cells["A1"].Value = "Filsti";
                    worksheet.Cells["B1"].Value = "URL kilde";
                    worksheet.Cells["C1"].Value = "Starttidspunkt (UTC)";
                    worksheet.Cells["D1"].Value = "Sluttidspunkt (UTC)";
                    worksheet.Cells["A1:D1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                        worksheet.Cells[i + 2, 4].Value = columns[3];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(1).Width = 80;
                    worksheet.Column(1).Style.WrapText = true;
                    worksheet.Column(2).Width = 120;
                    worksheet.Column(2).Style.WrapText = true;
                    worksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    worksheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }
            }

            if (type.Equals("Web Data"))
            {
                if (db_table.Equals("Autofyldhistorik"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Autofyldhistorik");

                    worksheet.Cells["A1"].Value = "Navn";
                    worksheet.Cells["B1"].Value = "Værdi";
                    worksheet.Cells["C1"].Value = "Antal brug";
                    worksheet.Cells["D1"].Value = "Oprettet (UTC)";
                    worksheet.Cells["E1"].Value = "Senest brugt (UTC)";
                    worksheet.Cells["A1:E1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                        worksheet.Cells[i + 2, 4].Value = columns[3];
                        worksheet.Cells[i + 2, 5].Value = columns[4];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(1).Width = 50;
                    worksheet.Column(1).Style.WrapText = true;
                    worksheet.Column(2).Width = 100;
                    worksheet.Column(2).Style.WrapText = true;
                    worksheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    worksheet.Column(5).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }

                if (db_table.Equals("Autofyldprofiler"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Autofyldprofiler");

                    worksheet.Cells["A1"].Value = "Fulde navn";
                    worksheet.Cells["B1"].Value = "E-mail";
                    worksheet.Cells["C1"].Value = "Telefonnummer";
                    worksheet.Cells["D1"].Value = "Adresse";
                    worksheet.Cells["E1"].Value = "Postnummer";
                    worksheet.Cells["F1"].Value = "By";
                    worksheet.Cells["G1"].Value = "Senest brugt (UTC)";
                    worksheet.Cells["A1:G1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                        worksheet.Cells[i + 2, 4].Value = columns[3];
                        worksheet.Cells[i + 2, 5].Value = columns[4];
                        worksheet.Cells[i + 2, 6].Value = columns[5];
                        worksheet.Cells[i + 2, 7].Value = columns[6];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(7).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }
            }

            if (type.Equals("Login Data"))
            {
                if (db_table.Equals("Logins"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Logins");

                    worksheet.Cells["A1"].Value = "URL";
                    worksheet.Cells["B1"].Value = "Brugernavn";
                    worksheet.Cells["C1"].Value = "Oprettet (UTC)";
                    worksheet.Cells["A1:C1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(1).Width = 100;
                    worksheet.Column(1).Style.WrapText = true;
                    worksheet.Column(2).Width = 100;
                    worksheet.Column(2).Style.WrapText = true;
                    worksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }
            }

            if (type.Equals("Shortcuts"))
            {
                if (db_table.Equals("Shortcuts"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Shortcuts");

                    worksheet.Cells["A1"].Value = "Indtastning";
                    worksheet.Cells["B1"].Value = "URL";
                    worksheet.Cells["C1"].Value = "Websidetitel";
                    worksheet.Cells["D1"].Value = "Senest besøgt (UTC)";
                    worksheet.Cells["A1:D1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                        worksheet.Cells[i + 2, 4].Value = columns[3];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(1).Width = 80;
                    worksheet.Column(1).Style.WrapText = true;
                    worksheet.Column(2).Width = 120;
                    worksheet.Column(2).Style.WrapText = true;
                    worksheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }
            }

            if (type.Equals("SyncData"))
            {
                if (db_table.Equals("Synchistorik"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Synchistorik");

                    worksheet.Cells["A1"].Value = "ID";
                    worksheet.Cells["B1"].Value = "Navn";
                    worksheet.Cells["C1"].Value = "Parsed indhold";
                    worksheet.Cells["D1"].Value = "Senest besøgt (UTC)";
                    worksheet.Cells["A1:D1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                        worksheet.Cells[i + 2, 4].Value = columns[3];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(3).Width = 150;
                    worksheet.Column(3).Style.WrapText = true;
                    worksheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }
            }

            if (type.Equals("Firefox"))
            {
                if (db_table.Equals("Webhistorik"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Webhistorik seneste");

                    worksheet.Cells["A1"].Value = "URL";
                    worksheet.Cells["B1"].Value = "Titel";
                    worksheet.Cells["C1"].Value = "Senest besøgt (UTC)";
                    worksheet.Cells["A1:C1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(1).Width = 100;
                    worksheet.Column(1).Style.WrapText = true;
                    worksheet.Column(2).Width = 100;
                    worksheet.Column(2).Style.WrapText = true;
                    worksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                    if (second_list.Count != 0)
                    {
                        ExcelWorksheet worksheet2 = excelpackage.Workbook.Worksheets.Add("Webhistorik alle");

                        worksheet2.Cells["A1"].Value = "URL";
                        worksheet2.Cells["B1"].Value = "Titel";
                        worksheet2.Cells["C1"].Value = "Senest besøgt (UTC)";
                        worksheet2.Cells["A1:C1"].Style.Font.Bold = true;

                        for (int i = 0; i < second_list.Count; i++)
                        {
                            string[] columns = second_list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                            worksheet2.Cells[i + 2, 1].Value = columns[0];
                            worksheet2.Cells[i + 2, 2].Value = columns[1];
                            worksheet2.Cells[i + 2, 3].Value = columns[2];
                        }

                        worksheet2.Cells[worksheet.Dimension.Address].AutoFitColumns();
                        worksheet2.Column(1).Width = 100;
                        worksheet2.Column(1).Style.WrapText = true;
                        worksheet2.Column(2).Width = 100;
                        worksheet2.Column(2).Style.WrapText = true;
                        worksheet2.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                        second_list.Clear();
                    }
                }

                if (db_table.Equals("Downloads"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Downloads");

                    worksheet.Cells["A1"].Value = "Filsti";
                    worksheet.Cells["B1"].Value = "URL kilde";
                    worksheet.Cells["C1"].Value = "Starttidspunkt (UTC)";
                    worksheet.Cells["D1"].Value = "Sluttidspunkt (UTC)";
                    worksheet.Cells["A1:D1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                        worksheet.Cells[i + 2, 4].Value = columns[3];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(1).Width = 80;
                    worksheet.Column(1).Style.WrapText = true;
                    worksheet.Column(2).Width = 120;
                    worksheet.Column(2).Style.WrapText = true;
                    worksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    worksheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }

                if (db_table.Equals("Formhistorik"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Formhistorik");

                    worksheet.Cells["A1"].Value = "Indtastning";
                    worksheet.Cells["B1"].Value = "Værdi";
                    worksheet.Cells["C1"].Value = "Senest brugt (UTC)";
                    worksheet.Cells["A1:C1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(1).Width = 80;
                    worksheet.Column(1).Style.WrapText = true;
                    worksheet.Column(2).Width = 120;
                    worksheet.Column(2).Style.WrapText = true;
                    worksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }

                if (db_table.Equals("Bookmarks"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Bookmarks");

                    worksheet.Cells["A1"].Value = "Bogmærketitel";
                    worksheet.Cells["B1"].Value = "URL";
                    worksheet.Cells["C1"].Value = "Tilføjet (UTC)";
                    worksheet.Cells["D1"].Value = "Senest ændret (UTC)";
                    worksheet.Cells["E1"].Value = "Senest besøgt (UTC)";
                    worksheet.Cells["A1:E1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                        worksheet.Cells[i + 2, 4].Value = columns[3];
                        worksheet.Cells[i + 2, 5].Value = columns[4];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(1).Width = 80;
                    worksheet.Column(1).Style.WrapText = true;
                    worksheet.Column(2).Width = 120;
                    worksheet.Column(2).Style.WrapText = true;
                    worksheet.Column(3).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    worksheet.Column(4).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    worksheet.Column(5).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }
            }

            if (type.Equals("Skype_new"))
            {
                ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Beskeder");

                worksheet.Cells["A1"].Value = "Samtale ID";
                worksheet.Cells["B1"].Value = "Type";
                worksheet.Cells["C1"].Value = "Tidspunkt (UTC)";
                worksheet.Cells["D1"].Value = "Oprettet af";
                worksheet.Cells["E1"].Value = "Indhold";
                worksheet.Cells["F1"].Value = "Egenskaber";
                worksheet.Cells["A1:F1"].Style.Font.Bold = true;

                // Parse kendte json tags (kan udvides når nye findes)
                for (int i = 0; i < list.Count; i++)
                {
                    dynamic json_data = JsonConvert.DeserializeObject(list[i]);

                    worksheet.Cells[i + 2, 1].Value = Convert.ToString(json_data.conversationId);
                    worksheet.Cells[i + 2, 2].Value = Convert.ToString(json_data.messagetype);
                    worksheet.Cells[i + 2, 3].Value = Convert.ToString(DateTimeOffset.FromUnixTimeMilliseconds(Convert.ToInt64(json_data.createdTime))).Substring(0, Convert.ToString(DateTimeOffset.FromUnixTimeMilliseconds(Convert.ToInt64(json_data.createdTime))).Length - 7);
                    worksheet.Cells[i + 2, 4].Value = Convert.ToString(json_data.creator);

                    if (json_data.messagetype == "Event/Call")
                    {
                        string json_data_content = json_data.content;

                        string[] split = json_data_content.Split(new string[] { "identity" }, StringSplitOptions.None);

                        string content = "";

                        if (split.Length > 1)
                            for (int j = 1; j < split.Length; j++)
                                content += "\r\n" + split[j].Substring(split[j].IndexOf("\""), split[j].LastIndexOf("\""));

                        worksheet.Cells[i + 2, 5].Value = "Deltagere:" + content;
                    }

                    else if (json_data.messagetype == "Notice")
                    {
                        JArray json_data_content = JsonConvert.DeserializeObject(Convert.ToString(json_data.content));
                        JToken content = json_data_content.First.SelectToken("attachments").First.SelectToken("content");

                        worksheet.Cells[i + 2, 5].Value = "Titel: " + content.SelectToken("title") +
                                                            "\r\nTekst: " + content.SelectToken("text");
                    }

                    else if (json_data.messagetype == "PopCard")
                    {
                        JArray json_data_content = JsonConvert.DeserializeObject(Convert.ToString(json_data.content));
                        JToken content = json_data_content.First.SelectToken("content");

                        worksheet.Cells[i + 2, 5].Value = "Titel: " + content.SelectToken("title") +
                                                            "\r\nTekst: " + content.SelectToken("text");
                    }

                    else if (json_data.messagetype == "ThreadActivity/AddMember" || json_data.messagetype == "ThreadActivity/DeleteMember" || json_data.messagetype == "ThreadActivity/TopicUpdate" || json_data.messagetype == "ThreadActivity/HistoryDisclosedUpdate")
                    {
                        string content = "";
                        XElement elements = XElement.Parse(Convert.ToString(json_data.content));

                        foreach (XElement element in elements.Descendants())
                        {
                            if (element.Name == "eventtime")
                                content += "Tidspunkt: " + Convert.ToString(DateTimeOffset.FromUnixTimeMilliseconds(Convert.ToInt64(element.Value))).Substring(0, Convert.ToString(DateTimeOffset.FromUnixTimeMilliseconds(Convert.ToInt64(element.Value))).Length - 7);
                            if (element.Name == "initiator")
                                content += "\r\nIgangsætter: " + element.Value;
                            if (element.Name == "target")
                                content += "\r\nMål: " + element.Value;
                            if (element.Name == "value")
                                content += "\r\nVærdi: " + element.Value;
                        }

                        worksheet.Cells[i + 2, 5].Value = content;
                    }

                    else if (Regex.IsMatch(Convert.ToString(json_data.messagetype), "RichText/"))
                    {
                        if (json_data.messagetype == "RichText/Sms")
                        {
                            XElement elements = XElement.Parse(Convert.ToString(json_data.content));

                            foreach (XElement element in elements.Descendants())
                                if (element.Name == "body")
                                    worksheet.Cells[i + 2, 5].Value = element.Value;
                        }
                        else
                        {
                            string filtered = Convert.ToString(json_data.content);
                            filtered = filtered.Substring(0, filtered.LastIndexOf("</URIObject>") + 12);

                            XElement elements = XElement.Parse(filtered);

                            worksheet.Cells[i + 2, 5].Value = elements.Value;
                        }
                    }

                    else
                        worksheet.Cells[i + 2, 5].Value = Convert.ToString(json_data.content);

                    if (json_data.messagetype == "Text")
                    {
                        if (Convert.ToString(json_data).Contains("properties") && Convert.ToString(json_data.properties).Contains("call-log"))
                        {
                            dynamic json_data_properties = JsonConvert.DeserializeObject(Convert.ToString(json_data.properties["call-log"]));

                            worksheet.Cells[i + 2, 6].Value = "Opkaldslog\r\nStarttidspunkt: " + Convert.ToString(json_data_properties.startTime) +
                                                                "\r\nSluttidspunkt: " + Convert.ToString(json_data_properties.endTime) +
                                                                "\r\nAfsender: " + Convert.ToString(json_data_properties.originator) +
                                                                "\r\nModtager: " + Convert.ToString(json_data_properties.target);
                        }

                        else
                            worksheet.Cells[i + 2, 6].Value = Convert.ToString(json_data.properties);
                    }
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                worksheet.Column(5).Width = 100;
                worksheet.Column(5).Style.WrapText = true;
                worksheet.Column(6).Width = 40;
                worksheet.Column(6).Style.WrapText = true;
            }

            if (type.Equals("Skype_old"))
            {
                if (db_table.Equals("Beskeder"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Beskeder");

                    worksheet.Cells["A1"].Value = "Samtale ID";
                    worksheet.Cells["B1"].Value = "Afsender navn";
                    worksheet.Cells["C1"].Value = "Afsender profilnavn";
                    worksheet.Cells["D1"].Value = "Modtager";
                    worksheet.Cells["E1"].Value = "Indhold";
                    worksheet.Cells["F1"].Value = "Tidspunkt (UTC)";
                    worksheet.Cells["A1:F1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                        worksheet.Cells[i + 2, 4].Value = columns[3];

                        string[] split = columns[4].Split(new string[] { "identity" }, StringSplitOptions.None);

                        string content = "";

                        if (split.Length > 1)
                        {
                            for (int j = 1; j < split.Length; j++)
                                content += "\r\n" + split[j].Substring(split[j].IndexOf("\""), split[j].LastIndexOf("\""));

                            worksheet.Cells[i + 2, 5].Value = "Deltagere:" + content;
                        }

                        else
                            worksheet.Cells[i + 2, 5].Value = columns[4];

                        worksheet.Cells[i + 2, 6].Value = columns[5];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(5).Width = 100;
                    worksheet.Column(5).Style.WrapText = true;
                    worksheet.Column(6).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                }

                if (db_table.Equals("Konti"))
                {
                    ExcelWorksheet worksheet = excelpackage.Workbook.Worksheets.Add("Konti");

                    worksheet.Cells["A1"].Value = "Skypenavn";
                    worksheet.Cells["B1"].Value = "Fulde navn";
                    worksheet.Cells["C1"].Value = "Fødselsdato";
                    worksheet.Cells["D1"].Value = "By";
                    worksheet.Cells["E1"].Value = "Telefonnummer";
                    worksheet.Cells["F1"].Value = "E-mail";
                    worksheet.Cells["G1"].Value = "Webside";
                    worksheet.Cells["A1:G1"].Style.Font.Bold = true;

                    for (int i = 0; i < list.Count; i++)
                    {
                        string[] columns = list[i].Split(new string[] { ";;" }, StringSplitOptions.None);

                        worksheet.Cells[i + 2, 1].Value = columns[0];
                        worksheet.Cells[i + 2, 2].Value = columns[1];
                        worksheet.Cells[i + 2, 3].Value = columns[2];
                        worksheet.Cells[i + 2, 4].Value = columns[3];
                        worksheet.Cells[i + 2, 5].Value = columns[4];
                        worksheet.Cells[i + 2, 6].Value = columns[5];
                        worksheet.Cells[i + 2, 7].Value = columns[6];
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }
            }

            string text_assnr = "";
            string text_kosternr = "";

            if (assnr.Text != "")
                text_assnr = assnr.Text + "_";

            if (kosternr.Text != "")
                text_kosternr = "koster_" + kosternr.Text + "_";

            // Eksporter til excel fil
            FileInfo excelfile = new FileInfo(new FileInfo(db_filename).DirectoryName + text_assnr + text_kosternr + new FileInfo(db_filename).Name + "_" + db_table + ".xlsx");

            try
            {
                excelpackage.SaveAs(excelfile);
                notify.Content = "Eksporteret til:";
                path.Text = excelfile.ToString();
                stackpanel_query.Visibility = Visibility.Visible;
            }
            catch (InvalidOperationException)
            {
                notify.Content = "Filen eksisterer og er allerede åben.";
                path.Text = excelfile.ToString();
                stackpanel_query.Visibility = Visibility.Hidden;
            }
        }

        // Parse tabel til liste
        private List<string> ParseTableToList(string db, string table)
        {
            List<string> list = new List<string>();
            string type = db_type.Split(',')[0];
            db_table = table;

            SQLiteConnection connect = new SQLiteConnection("Data Source=" + db);
            connect.Open();
            
            SQLiteCommand command = connect.CreateCommand();

            if (type.Equals("Skype_new"))
            {
                if (table.Equals("Beskeder"))
                {
                    command.CommandText = query.Text = "SELECT nsp_data FROM messagesv12";
                    SQLiteDataReader reader = command.ExecuteReader();
                    
                    while (reader.Read())
                        list.Add(Convert.ToString(reader["nsp_data"]));
                }
            }

            if (type.Equals("Skype_old"))
            {
                if (table.Equals("Beskeder"))
                {
                    command.CommandText = query.Text = "SELECT id, author, from_dispname, dialog_partner, body_xml, datetime(timestamp, 'unixepoch') FROM Messages";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)) + ";;" + Convert.ToString(reader.GetValue(3)) + ";;" + Convert.ToString(reader.GetValue(4)) + ";;" + Convert.ToString(reader.GetValue(5)));
                }

                if (table.Equals("Konti"))
                {
                    command.CommandText = query.Text = "SELECT skypename, fullname, birthday, city, phone_mobile, emails, homepage FROM Accounts";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)) + ";;" + Convert.ToString(reader.GetValue(3)) + ";;" + Convert.ToString(reader.GetValue(4)) + ";;" + Convert.ToString(reader.GetValue(5)) + ";;" + Convert.ToString(reader.GetValue(6)));
                }
            }

            if (type.Equals("History"))
            {
                if (table.Equals("Webhistorik"))
                {
                    command.CommandText = query.Text = "SELECT url, title, datetime(last_visit_time / 1000000 + (strftime('%s', '1601-01-01')), 'unixepoch') FROM urls";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)));

                    reader.Close();

                    SQLiteCommand command2 = connect.CreateCommand();
                    command2.CommandText = "SELECT urls.url, title, datetime(visit_time / 1000000 + (strftime('%s', '1601-01-01')), 'unixepoch') FROM urls INNER JOIN visits ON urls.id = visits.url ORDER BY visits.visit_time asc";
                    reader = command2.ExecuteReader();

                    query.Text = command.CommandText + "\r\n\r\nog\r\n\r\n" + command2.CommandText;

                    while (reader.Read())
                        second_list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)));
                }

                if (table.Equals("Downloads"))
                {
                    command.CommandText = query.Text = "SELECT downloads.target_path, downloads_url_chains.url, datetime(start_time / 1000000 + (strftime('%s', '1601-01-01')), 'unixepoch'), datetime(end_time / 1000000 + (strftime('%s', '1601-01-01')), 'unixepoch') FROM downloads INNER JOIN downloads_url_chains ON downloads.id = downloads_url_chains.id GROUP BY downloads.id";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)) + ";;" + Convert.ToString(reader.GetValue(3)));
                }
            }

            if (type.Equals("Web Data"))
            {
                if (table.Equals("Autofyldhistorik"))
                {
                    command.CommandText = query.Text = "SELECT autofill.name, autofill.value, autofill.count, datetime(date_created, 'unixepoch'), datetime(date_last_used, 'unixepoch') FROM autofill";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)) + ";;" + Convert.ToString(reader.GetValue(3)) + ";;" + Convert.ToString(reader.GetValue(4)));
                }

                if (table.Equals("Autofyldprofiler"))
                {
                    command.CommandText = query.Text = "SELECT autofill_profile_names.full_name, autofill_profile_emails.email, autofill_profile_phones.number, autofill_profiles.street_address, autofill_profiles.zipcode, autofill_profiles.city, datetime(use_date, 'unixepoch') FROM autofill_profiles INNER JOIN autofill_profile_emails ON autofill_profiles.guid = autofill_profile_emails.guid INNER JOIN autofill_profile_names ON autofill_profiles.guid = autofill_profile_names.guid INNER JOIN autofill_profile_phones ON autofill_profiles.guid = autofill_profile_phones.guid";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)) + ";;" + Convert.ToString(reader.GetValue(3)) + ";;" + Convert.ToString(reader.GetValue(4)) + ";;" + Convert.ToString(reader.GetValue(5)) + ";;" + Convert.ToString(reader.GetValue(6)));
                }
            }

            if (type.Equals("Login Data"))
            {
                if (table.Equals("Logins"))
                {
                    command.CommandText = query.Text = "SELECT origin_url, username_value, datetime(date_created / 1000000 + (strftime('%s', '1601-01-01')), 'unixepoch') FROM logins";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)));
                }
            }

            if (type.Equals("Shortcuts"))
            {
                if (table.Equals("Shortcuts"))
                {
                    command.CommandText = query.Text = "SELECT fill_into_edit, url, description, datetime(last_access_time / 1000000 + (strftime('%s', '1601-01-01')), 'unixepoch') FROM omni_box_shortcuts";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)) + ";;" + Convert.ToString(reader.GetValue(3)));
                }
            }

            if (type.Equals("SyncData"))
            {
                if (table.Equals("Synchistorik"))
                {
                    command.CommandText = query.Text = "SELECT metahandle, non_unique_name, hex(specifics), datetime(mtime / 1000, 'unixepoch') FROM metas";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)) + ";;" + Convert.ToString(reader.GetValue(3)));
                }
            }

            if (type.Equals("Firefox"))
            {
                if (table.Equals("Webhistorik"))
                {
                    command.CommandText = "SELECT url, title, datetime(last_visit_date / 1000000, 'unixepoch') FROM moz_places";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)));

                    reader.Close();

                    SQLiteCommand command2 = connect.CreateCommand();
                    command2.CommandText = "SELECT url, title, datetime(visit_date / 1000000, 'unixepoch') FROM moz_places INNER JOIN moz_historyvisits ON moz_places.id = moz_historyvisits.place_id ORDER BY moz_historyvisits.visit_date asc";
                    reader = command2.ExecuteReader();

                    query.Text = command.CommandText + "\r\n\r\nog\r\n\r\n" + command2.CommandText;

                    while (reader.Read())
                        second_list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)));
                }

                if (table.Equals("Downloads"))
                {
                    command.CommandText = query.Text = "SELECT content, moz_places.url, datetime(dateAdded / 1000000, 'unixepoch'), datetime(lastModified / 1000000, 'unixepoch') FROM moz_annos INNER JOIN moz_places ON moz_annos.place_id = moz_places.id GROUP BY moz_annos.place_id";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)) + ";;" + Convert.ToString(reader.GetValue(3)));
                }

                if (table.Equals("Formhistorik"))
                {
                    command.CommandText = query.Text = "SELECT fieldname, value, datetime(lastUsed / 1000000, 'unixepoch') FROM moz_formhistory";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)));
                }

                if (table.Equals("Bookmarks"))
                {
                    command.CommandText = query.Text = "SELECT moz_bookmarks.title, moz_places.url, datetime(moz_bookmarks.dateAdded / 1000000, 'unixepoch'), datetime(moz_bookmarks.lastModified / 1000000, 'unixepoch'), datetime(moz_places.last_visit_date / 1000000, 'unixepoch') FROM moz_places INNER JOIN moz_bookmarks ON moz_places.id = moz_bookmarks.fk ORDER BY moz_bookmarks.dateAdded ASC";
                    SQLiteDataReader reader = command.ExecuteReader();

                    while (reader.Read())
                        list.Add(Convert.ToString(reader.GetValue(0)) + ";;" + Convert.ToString(reader.GetValue(1)) + ";;" + Convert.ToString(reader.GetValue(2)) + ";;" + Convert.ToString(reader.GetValue(3)) + ";;" + Convert.ToString(reader.GetValue(4)));
                }
            }

            connect.Close();
            return list;
        }

        // Analysér filnavn og tabelnavne
        private string DetermineDatabase(string file)
        {
            string table_list = "";

            SQLiteConnection connect = new SQLiteConnection("Data Source=" + file);
            connect.Open();

            SQLiteCommand command = connect.CreateCommand();
            command.CommandText = "SELECT name FROM sqlite_master WHERE type='table'";

            try
            {
                SQLiteDataReader reader = command.ExecuteReader();

                while (reader.Read())
                    table_list += reader["name"] + " ";

                connect.Close();

                if (new FileInfo(file).Name.Equals("History") && table_list.Contains("downloads") && table_list.Contains("urls"))
                    return "History,Downloads,Webhistorik";
                
                if (new FileInfo(file).Name.Equals("Web Data") && table_list.Contains("autofill") && table_list.Contains("autofill_profiles"))
                    return "Web Data,Autofyldhistorik,Autofyldprofiler";
                
                if (new FileInfo(file).Name.Equals("Login Data") && table_list.Contains("logins"))
                    return "Login Data,Logins";

                if (new FileInfo(file).Name.Equals("Shortcuts") && table_list.Contains("omni_box_shortcuts"))
                    return "Shortcuts,Shortcuts";

                if (new FileInfo(file).Name.Equals("SyncData.sqlite3") && table_list.Contains("metas"))
                    return "SyncData,Synchistorik";

                if (new FileInfo(file).Name.Equals("places.sqlite") && table_list.Contains("moz_annos") && table_list.Contains("moz_places") && table_list.Contains("moz_bookmarks"))
                    return "Firefox,Downloads,Webhistorik,Bookmarks";

                if (new FileInfo(file).Name.Equals("formhistory.sqlite") && table_list.Contains("moz_formhistory"))
                    return "Firefox,Formhistorik";

                if (new FileInfo(file).Extension.Equals(".db") && table_list.Contains("messagesv12"))
                    return "Skype_new,Beskeder";
                
                if (new FileInfo(file).Name.Equals("main.db") && table_list.Contains("Messages") && table_list.Contains("Accounts"))
                    return "Skype_old,Beskeder,Konti";
                
                else
                    return "Unknown";
            }
            catch (SQLiteException)
            {
                return "Error";
            }
        }

        // Åbn menu
        private void open_click(object sender, RoutedEventArgs e)
        {
            choose.Content = "";
            notify.Content = "";
            path.Clear();
            stackpanel_buttons.Children.Clear();
            stackpanel_query.Visibility = Visibility.Hidden;

            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            bool? result = dialog.ShowDialog();

            if (result == true)
            {
                db_type = DetermineDatabase(dialog.FileName);

                if (db_type.Equals("Unknown"))
                    choose.Content = "Ukendt database";
                else if (db_type.Equals("Error"))
                    choose.Content = "Ingen database";
                else
                {
                    string[] split = db_type.Split(',');
                    choose.Content = "Vælg hvad, der skal eksporteres";

                    for (int i = 1; i < split.Length; i++)
                    {
                        Button button = new Button();
                        button.Foreground = Brushes.LightSteelBlue;
                        button.Background = Brushes.MidnightBlue;
                        button.BorderBrush = Brushes.LightSteelBlue;
                        button.Margin = new Thickness(5, 0, 0, 0);
                        button.Name = split[i];
                        button.Content = " " + split[i] + " ";
                        button.Click += table_click;
                        buttons_list.Add(button);
                        stackpanel_buttons.Children.Add(button);
                    }

                    db_filename = new FileInfo(dialog.FileName).FullName;
                }
            }
        }

        private void table_click(object sender, RoutedEventArgs e)
        {
            Button source = (Button)e.Source;
            table_list = ParseTableToList(db_filename, source.Name);
            ExportToExcel(table_list);
        }

        private void dragenter(object sender, DragEventArgs e)
        {
            dragdrop_grid.Visibility = Visibility.Visible;
            dragdrop_grid.Opacity = 0.7;
        }

        private void dragleave(object sender, DragEventArgs e)
        {
            dragdrop_grid.Visibility = Visibility.Hidden;
            dragdrop_grid.Opacity = 0;
        }

        private void clipboard_click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(query.Text);
        }

        private void supported_mouseenter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            OpacityAnimation(supported_text, supported_text.Opacity, 1, 0.5);
            HeighAnimation(supported_text, supported_text.Height, 450, 0.5);
        }

        private void supported_mouseleave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            OpacityAnimation(supported_text, supported_text.Opacity, 0, 0.5);
            HeighAnimation(supported_text, supported_text.Height, 0, 0.5);
        }

        // Drag n drop funktion
        private void dragdrop(object sender, DragEventArgs e)
        {
            dragdrop_grid.Visibility = Visibility.Hidden;
            dragdrop_grid.Opacity = 0;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                choose.Content = "";
                notify.Content = "";
                path.Clear();
                stackpanel_buttons.Children.Clear();
                stackpanel_query.Visibility = Visibility.Hidden;

                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                db_type = DetermineDatabase(files[0]);

                if (db_type.Equals("Unknown"))
                    choose.Content = "Ukendt database";
                else if (db_type.Equals("Error"))
                    choose.Content = "Ingen database";
                else
                {
                    string[] split = db_type.Split(',');
                    choose.Content = "Vælg hvad, der skal eksporteres";

                    for (int i = 1; i < split.Length; i++)
                    {
                        Button button = new Button();
                        button.Foreground = Brushes.LightSteelBlue;
                        button.Background = Brushes.MidnightBlue;
                        button.BorderBrush = Brushes.LightSteelBlue;
                        button.Margin = new Thickness(5, 0, 0, 0);
                        button.Name = split[i];
                        button.Content = " " + split[i] + " ";
                        button.Click += table_click;
                        buttons_list.Add(button);
                        stackpanel_buttons.Children.Add(button);
                    }

                    db_filename = new FileInfo(files[0]).FullName;
                }
            }
        }

        private void OpacityAnimation(UIElement control, double from, double to, double duration)
        {
            CircleEase ce = new CircleEase();
            ce.EasingMode = EasingMode.EaseOut;
            DoubleAnimation da = new DoubleAnimation();
            da.From = from;
            da.To = to;
            da.Duration = TimeSpan.FromSeconds(duration);
            da.EasingFunction = ce;
            control.BeginAnimation(OpacityProperty, da);
        }

        private void HeighAnimation( UIElement control, double from, double to, double duration)
        {
            CircleEase ce = new CircleEase();
            ce.EasingMode = EasingMode.EaseOut;
            DoubleAnimation da = new DoubleAnimation();
            da.From = from;
            da.To = to;
            da.Duration = TimeSpan.FromSeconds(duration);
            da.EasingFunction = ce;
            control.BeginAnimation(HeightProperty, da);
        }
    }
}
