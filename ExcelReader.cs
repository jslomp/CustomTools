/*
 * Author: Jacob Slomp
 */


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.Collections.ObjectModel;
using System.IO;

namespace CustomTools
{
    class ExcelReader
    {

        ~ExcelReader()
        {
            try
            {
                // destrector
                xlApp.Quit();
                // Your code
            }
            catch (InvalidCastException e)
            {
                Console.WriteLine(e.Message);
            }
        }


        string filePath = "";
        List<string> headers = new List<string>();
        private int counter = 2;

        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;

        string str;
        //int rCnt;
        //int cCnt;
        int rw = 0;
        int cl = 0;
        //int TotalSheets = 0;

        public Boolean keep_columns = false;



        DataGridView myGrid = null;

        public ExcelReader(string text = "")
        {
            if (text != "" && text != null)
            {
                setFile(text);
            }
        }
        public void resetCount()
        {
            counter = 2;
        }

        public void openFileFromDialog()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                //openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel (*.xls)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    setFile(openFileDialog.FileName);


                }
            }

        }
        public string delimiter = ",";
        public int ExcelFormat = 5;
        public void setDelimiter(string delimiter)
        {
            this.delimiter = delimiter;
            if(delimiter == ",") { 
                this.ExcelFormat = 2;
            } else if (delimiter == ";")
            {
                this.ExcelFormat = 4;
            } else if (delimiter == "\t")
            {
                this.ExcelFormat = 1;
            } else
            {
                this.ExcelFormat = 6;
            }
        }
        public void setFile(string path)
        {
            if (File.Exists(path))
            {

                filePath = Path.GetFullPath(path);

                if (path.EndsWith(".csv"))
                {
                    this.ExcelFormat = 6;
                }
                xlApp = new Excel.Application();
                //workbooks.Open(Filename, [UpdateLinks], [ReadOnly], [Format], [Password], [WriteResPassword], [IgnoreReadOnlyRecommended], [Origin], **[Delimiter]**, [Editable], [Notify], [Converter], [AddToMru], [Local], [CorruptLoad]) 
                xlWorkBook = xlApp.Workbooks.Open(filePath, 0, true, this.ExcelFormat, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, delimiter, false, false, 0, true, 1, 0);


                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                range = xlWorkSheet.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;


                int colcount = 0;
                headers = new List<string>();
                for (int cols = 1; cols <= cl; cols++)
                {
                    string str = (string)(range.Cells[1, cols] as Excel.Range).Value;
                    if (str != null)
                    {
                        str = str.Replace("\n", " ");
                        str = str.Replace("\t", " ");
                        str = str.Replace("\r", " ");

                        headers.Add(str);
                        colcount = 0;
                    }
                    else
                    {
                        str = "";
                        headers.Add(str);
                        colcount++;
                    }
                    if(colcount > 3)
                    {
                        cl = cols;
                        break;
                    }
                }
            }
            else
            {
                MessageBox.Show("Cant find file...");
            }


        }

        public List<string> getHeaders()
        {
            if (filePath == "")
            {
                return null;
            }
            return headers;

        }


        public void ClearGrid()
        {
            if (myGrid != null)
            {
                myGrid.Rows.Clear();
                if (keep_columns == false)
                {
                    myGrid.Columns.Clear();
                }
            }

        }
        public void attachToGrid(DataGridView d)
        {
            myGrid = d;
        }

        int hasValues = 0;
        public List<string> getRawRow()
        {


            if (filePath == "")
            {
                return null;
            }
            List<string> values = new List<string>();
            Boolean hasValue = false;

            //str = (string)Convert.ToString((range.Cells[counter, 1] as Excel.Range).Value);
            //if (str != null && str.Length > 0)
            //{

            for (int cols = 1; cols <= cl; cols++)
            {
                str = (string)Convert.ToString((range.Cells[counter, cols] as Excel.Range).Value);
                values.Add(str);

                if (str != null && str != "")
                {
                    hasValue = true;
                }
            }

            // }

            counter++;
            if (hasValue)
            {
                hasValues = 0;
                return values;

            }
            else
            {
                hasValues++;
                if (hasValues > 3)
                {
                    xlWorkBook.Close();
                    xlApp.Quit();
                    return null;
                }
                else
                {
                    return getRawRow();
                }
            }

        }


        public Dictionary<string, string> getRow()
        {
            if (filePath == "")
            {
                return null;
            }
            Dictionary<string, string> values = new Dictionary<string, string>();
            Boolean hasValue = false;

            str = (string)Convert.ToString((range.Cells[counter, 1] as Excel.Range).Value);
            if (str != null && str.Length > 0)
            {

                for (int cols = 1; cols <= cl; cols++)
                {
                    str = (string)Convert.ToString((range.Cells[counter, cols] as Excel.Range).Value);

                    values.Add(headers[cols - 1], str);


                    if (str != "" && str != null)
                    {
                        hasValue = true;
                    }
                }



            }
            //Console.WriteLine(filePath+" | Line: "+counter+" has value: "+hasValue.ToString()+" | Cols: "+cl.ToString()+"");

            counter++;
            if (hasValue)
            {
                return values;

            }
            else
            {
                if (counter <= rw)
                {
                    return getRow();
                }
                else
                {
                    return null;
                }
            }

        }

        public void PushToGrid()
        {

            ClearGrid();
            if (keep_columns == false)
            {
                foreach (string h in headers)
                {
                    myGrid.Columns.Add(h, h);
                }
            }

            List<string> colnames = new List<string>();
            for (int i = 0; i < myGrid.Columns.Count; i++)
            {
                colnames.Add(myGrid.Columns[i].Name.ToLower());
            }

            Dictionary<string, string> data = getRow();
            while (data != null)
            {
                int index = myGrid.Rows.Add();
                foreach (KeyValuePair<string, string> cols in data)
                {

                    int ind = colnames.IndexOf(cols.Key.ToLower());
                    if (ind > -1)
                    {
                        myGrid.Rows[index].Cells[ind].Value = cols.Value;
                    }
                }
                data = getRow();

            }
        }
    }
}
