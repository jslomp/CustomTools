/*
 * Author: Jacob Slomp
 */


using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace hexagon_v2
{
    class ExcelExport
    {
        private List<Dictionary<string, string>> result2 = new List<Dictionary<string, string>>();
        private string[] header;

        public ExcelExport(List<Dictionary<string, string>> result2 = null)
        {
            this.result2 = result2;
            if(this.result2 == null)
            {
                this.result2 = new List<Dictionary<string, string>>();
            }
        }

        public void setHeader(List<string> header)
        {
            this.header = header.ToArray();
        }
        public void addRow(Dictionary<string, string> row)
        {
            this.result2.Add(row);
        }

        public string escape(string text)
        {
            if(text == null)
            {
                return "";
            }
            text = text.Replace("&", "&amp;");
            text = text.Replace("<", "&lt;");
            text = text.Replace(">", "&gt;");
            
            return text;
        }
        public string escapeCsv(string text)
        {
            if (text == null)
            {
                text = "";
            }
            
            text = text.Replace("\"", "\"\"");
            return text;
        }
        public string generateHTMLString()
        {

            int TotalColumns = header.Length;
            

            string HTML = "<!DOCTYPE HTML><html><head><style>body,html { font-size: 12px; font-family: calibri, arial, verdana, helvetica; margin:0; padding: 0; } table {border-collapse: collapse; width: 100%; } td { border:1px solid #000; padding: 5px; } span.v {writing-mode: vertical-rl;  text-orientation: upright; } </style></head><body><table style='width:100%'>";
         
            HTML += "<tr>";
            for (int i = 0; i < header.Length; i++)
            {
                string name = header[i];
                HTML += "<td>"+name+"</td>";
            }
            HTML += "</tr>";
            
            int line = 0;

            foreach (var res in result2)
            {
                line++;

                HTML += "<tr>";
                for (int i = 0; i < header.Length; i++)
                {
                    string name = header[i];

                    
                    if (res.ContainsKey(name))
                    {
                        HTML += "<td>" + res[name] + "</td>";
                        
                    }
                    else
                    {

                    }
                    
                    
                }
                HTML += "</tr>";

            }
            HTML += "</table></body></html>";
            return HTML;
        }
        public void generateHTML(string FileName)
        {
            File.WriteAllText(FileName, generateHTMLString());
        }
        public void generateCsv(string FileName)
        {
            
            try
            {
                File.WriteAllText(FileName, "");
            }catch(IOException e)
            {
                MessageBox.Show(e.Message);
                File.WriteAllText(FileName, "");
            }
            int TotalColumns = header.Length;

            
            for (int i=0; i < header.Length; i++)
            {
                string name = header[i];
                File.AppendAllText(FileName, "\""+ escapeCsv(name)+ "\"");
                if (i < header.Length-1)
                {
                    File.AppendAllText(FileName, ",");
                }
            }

            
            File.AppendAllText(FileName, "\n");
            int line = 0;

            foreach (var res in result2)
            {
                line++;
                
                for (int i = 0; i < header.Length; i++)
                {
                    string name = header[i];
                    if (res.ContainsKey(name))
                    {
                        
                        File.AppendAllText(FileName, "\""+ escapeCsv(res[name])+"\"");
                    }
                    
                    if (i < header.Length - 1)
                    {
                        File.AppendAllText(FileName, ",");
                    }
                }
                File.AppendAllText(FileName, "\n");

            }

        }
        public void generateXlsx(string FileName)
        {
            
            if (Directory.Exists("data"))
            {
                Directory.Delete("data",true);
            }
            Directory.CreateDirectory("data");
            if (!Directory.Exists("data/_rels"))
            {
                Directory.CreateDirectory("data/_rels");
            }
            if (!Directory.Exists("data/xl"))
            {
                Directory.CreateDirectory("data/xl");
            }
            if (!Directory.Exists("data/xl/_rels"))
            {
                Directory.CreateDirectory("data/xl/_rels");
            }
            if (!Directory.Exists("data/xl/worksheets"))
            {
                Directory.CreateDirectory("data/xl/worksheets");
            }
            List<string> files = new List<string>();
            File.WriteAllText("data/[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"utf-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" /><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" /><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" /><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" /></Types>");
            files.Add(@"data\[Content_Types].xml");
            
            File.WriteAllText("data/_rels/.rels", "<?xml version=\"1.0\" encoding=\"utf-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\" Id=\"R9750878102eb43a8\" /></Relationships>");
            files.Add(@"data\_rels\.rels");

            File.WriteAllText("data/xl/_rels/workbook.xml.rels", "<?xml version=\"1.0\" encoding=\"utf-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"/xl/styles.xml\" Id=\"R243ab635dacf49a5\" /><Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/xl/worksheets/sheet1.xml\" Id=\"Re1b99366d7f041aa\" /></Relationships>");
            files.Add(@"data\xl\_rels\workbook.xml.rels");


            File.WriteAllText("data/xl/workbook.xml", "<?xml version=\"1.0\" encoding=\"utf-8\"?><x:workbook xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheets><x:sheet name=\"Exported from SC\" sheetId=\"1\" r:id=\"Re1b99366d7f041aa\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" /></x:sheets></x:workbook>");
            files.Add(@"data\xl\workbook.xml");


            File.WriteAllText("data/xl/styles.xml", "<?xml version=\"1.0\" encoding=\"utf-8\"?><x:styleSheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:fonts><x:font><x:sz val=\"11\" /><x:color rgb=\"000000\" /><x:name val=\"Calibri\" /></x:font><x:font><x:b /><x:sz val=\"11\" /><x:color rgb=\"000000\" /><x:name val=\"Calibri\" /></x:font><x:font><x:i /><x:sz val=\"11\" /><x:color rgb=\"000000\" /><x:name val=\"Calibri\" /></x:font></x:fonts><x:fills><x:fill><x:patternFill patternType=\"none\" /></x:fill><x:fill><x:patternFill patternType=\"gray125\" /></x:fill><x:fill><x:patternFill patternType=\"solid\"><x:fgColor rgb=\"FFFFFF00\" /></x:patternFill></x:fill><x:fill><x:patternFill patternType=\"solid\"><x:fgColor rgb=\"FF00FF00\" /></x:patternFill></x:fill><x:fill><x:patternFill patternType=\"solid\"><x:fgColor rgb=\"FFFF0000\" /></x:patternFill></x:fill><x:fill><x:patternFill patternType=\"solid\"><x:fgColor rgb=\"FFFFFFCC\" /></x:patternFill></x:fill><x:fill><x:patternFill patternType=\"solid\"><x:fgColor rgb=\"FF993366\" /></x:patternFill></x:fill></x:fills><x:borders><x:border><x:left /><x:right /><x:top /><x:bottom /><x:diagonal /></x:border><x:border><x:left style=\"thin\"><x:color auto=\"1\" /></x:left><x:right style=\"thin\"><x:color auto=\"1\" /></x:right><x:top style=\"thin\"><x:color auto=\"1\" /></x:top><x:bottom style=\"thin\"><x:color auto=\"1\" /></x:bottom><x:diagonal /></x:border></x:borders><x:cellXfs><x:xf fontId=\"0\" fillId=\"0\" borderId=\"0\" /><x:xf fontId=\"1\" fillId=\"5\" borderId=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" /><x:xf fontId=\"0\" fillId=\"0\" borderId=\"1\" applyBorder=\"1\" /><x:xf fontId=\"0\" fillId=\"2\" borderId=\"1\" applyFill=\"1\" applyBorder=\"1\" /><x:xf fontId=\"0\" fillId=\"3\" borderId=\"1\" applyFill=\"1\" applyBorder=\"1\" /><x:xf fontId=\"0\" fillId=\"4\" borderId=\"1\" applyFill=\"1\" applyBorder=\"1\" /><x:xf fontId=\"0\" fillId=\"6\" borderId=\"1\" applyFill=\"1\" applyBorder=\"1\" /><x:xf fontId=\"2\" fillId=\"0\" borderId=\"1\" applyFont=\"1\" applyBorder=\"1\" /><x:xf fontId=\"0\" fillId=\"0\" borderId=\"0\" applyAlignment=\"1\"><x:alignment horizontal=\"center\" vertical=\"center\" /></x:xf><x:xf numFmtId=\"22\" applyNumberFormat=\"1\" /><x:xf numFmtId=\"14\" applyNumberFormat=\"1\" /></x:cellXfs></x:styleSheet>");
            files.Add(@"data\xl\styles.xml");


            string head = "<?xml version=\"1.0\" encoding=\"utf-8\"?><x:worksheet xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><x:sheetData>";
            File.WriteAllText(@"data\xl\worksheets\sheet1.xml", head);
            files.Add(@"data\xl\worksheets\sheet1.xml");

            string row = "<x:row>";
            foreach (var name in header)
            {

                row += "<x:c t=\"inlineStr\" s=\"1\">";
                row += "<x:is><x:t>" + escape(name) + "</x:t></x:is>";
                row += "</x:c>";
            }
            row += "</x:row>";


            File.AppendAllText("data/xl/worksheets/sheet1.xml", row);


            int line = 0;
            foreach(var res in result2)
            {
                line++;

                row = "<x:row>";
                foreach (var name in header)
                {

                    row += "<x:c t=\"inlineStr\">";
                    if (res.ContainsKey(name))
                    {
                        row += "<x:is><x:t>" + escape(res[name]) + "</x:t></x:is>";
                    }else
                    {
                        row += "<x:is><x:t></x:t></x:is>";
                    }
                    row += "</x:c>";
                }
                row += "</x:row>";


                File.AppendAllText("data/xl/worksheets/sheet1.xml", row);
            }

            string footer = "</x:sheetData></x:worksheet>";
            File.AppendAllText("data/xl/worksheets/sheet1.xml", footer);


            createZip(FileName, files.ToArray());

        }

        private string makeHeader(string v)
        {
            
            return v;
        }


        
        public void createZip(string FileName, string[] files)
        {
            if (File.Exists(FileName))
            {
                try
                {
                    File.Delete(FileName);
                }catch(IOException e)
                {
                    MessageBox.Show(e.Message);
                    File.Delete(FileName);
                }
            }

            // Create FileStream for output ZIP archive
            using (var fileStream = new FileStream(FileName, FileMode.CreateNew))
            {
                using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Create, true))
                {

                    foreach (string f in files)
                    {
                        var fileBytes = File.ReadAllBytes(f);
                        string name = f;
                        name = name.Replace("data/", "/");
                        name = name.Replace(@"data\", @"");
                        var zipArchiveEntry = archive.CreateEntry(name, CompressionLevel.Optimal);
                        using (var zipStream = zipArchiveEntry.Open())
                        {
                            zipStream.Write(fileBytes, 0, fileBytes.Length);

                        }

                    }

                }
            }

        }

        public void saveAs(string FileName, string type = "")
        {
            if(FileName == null || FileName == "")
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel file xlsx | *.xlsx | CSV file | *.csv | HTML File | *.html";
                var result = sfd.ShowDialog();
                if(result == DialogResult.OK)
                {
                    FileName = sfd.FileName;
                }
            }
            if(FileName != null && FileName != "") { 
                if(FileName.ToLower().EndsWith(".xlsx"))
                {
                    generateXlsx(FileName);
                }
                if (FileName.ToLower().EndsWith(".csv"))
                {
                    generateCsv(FileName);
                }
                if (FileName.ToLower().EndsWith(".html"))
                {
                    generateHTML(FileName);
                }
            } else
            {
                MessageBox.Show("Cannot save "+FileName);
            }
        }

        internal void setHeader(string[] vAssets)
        {
            this.header = vAssets;
        }
    }
}
