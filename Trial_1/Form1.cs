using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using ExcelDataReader;
namespace Trial_1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataSet result;
        private void RDBtn_MouseClick(object sender, MouseEventArgs e)
        {
            List<string> NMedList;
            //Open dialog and choose raw files
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;
            dialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            dialog.Title = "Select a text file";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                int MedRow = 2; //keep tracking rows
                //read in each file
                foreach (string fname in dialog.FileNames)// loop through each selected file 
                {
                    string[] lines = System.IO.File.ReadAllLines(fname);
                    if (lines[0] == "ecwPtStatement,1.0")// this is kind a switch statement to determine the format for each clients
                    {
                        NMedList = System.IO.File.ReadAllLines(fname).ToList();
                        DateTime today = DateTime.Today;
                        string date = today.ToString("dd-MM-yyyy");
                        string[] dateSplit = date.Split('-');
                        string location = dateSplit[2] + @"\" + dateSplit[1] + @"\" + date;
                        string newFolder = @"C:\Invoice\Clients\Northern Medical Group\" + location;
                        string ExcelFolder = @"C:\Invoice\Excel Files";
                        //create a folder based on current year, month, and date
                        Directory.CreateDirectory(newFolder);
                        string comText = newFolder + @"\NMED01_" + dateSplit[2] + "_" + dateSplit[1] + "_" + date + ".txt";
                        string ExcFile = newFolder + @"\NMED01_" + dateSplit[2] + "_" + dateSplit[1] + "_" + date + ".xlsx";
                        string ExcFile2 = ExcelFolder + @"\NMED01_" + dateSplit[2] + "_" + dateSplit[1] + "_" + date + ".xlsx";
                        /*if (!System.IO.File.Exists(comText))
                        {
                            FileStream createtextFile = File.Create(comText);
                        }*/
                        File.WriteAllLines(comText, NMedList);
                        if (!System.IO.File.Exists(ExcFile))
                        {
                            ExcelPackage med = new ExcelPackage();
                            med.Workbook.Worksheets.Add("Northern Medical Group");//creat work sheet
                            var headerRow = new List<string[]>()//add header
                            {
                                new string[] { "ACCOUNT","ID", "First Name", "MIDDLE NAME", "Last Name", "ADDRESSLINE1", "ADDRESSLINE2", "CITY", "STATE", "ZIPCODE" }
                            };
                            string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                            var worksheet = med.Workbook.Worksheets["Northern Medical Group"]; // load worksheet based on name
                            worksheet.Cells[headerRange].LoadFromArrays(headerRow); //loead data to worksheet 
                            FileInfo excelFile = new FileInfo(ExcFile);
                            FileInfo excelFile2 = new FileInfo(ExcFile2);
                            med.SaveAs(excelFile); // save file
                            med.SaveAs(excelFile2);
                        }
                        FileInfo loadfile = new FileInfo(ExcFile);// load excel file
                        FileInfo copyFile = new FileInfo(ExcFile2);
                        ExcelPackage med1 = new ExcelPackage(loadfile);
                        var worksheet1 = med1.Workbook.Worksheets["Northern Medical Group"];
                        for (int i = 1; i < lines.Length; i++)
                        {// read in each line
                            if (lines[i] == "ecwPtStatement") // indicator for client format  
                            {
                                var j = i + 1;
                                if (j < lines.Length)
                                {
                                    string[] info = lines[j].Split(',');
                                    string guid = System.Guid.NewGuid().ToString().Replace("-", "").ToUpper();
                                    var data = new List<string[]>()
                                    {
                                            new string[]{info[4].Trim('"'), guid, info[0].Trim('"'), info[1].Trim('"'), info[2].Trim('"'), info[9].Trim('"'), info[10].Trim('"'), info[11].Trim('"'), info[12].Trim('"'), info[13].Trim('"') }
                                    };
                                    worksheet1.Cells[MedRow, 1].LoadFromArrays(data);
                                    MedRow++;
                                }
                            }
                        }
                        med1.SaveAs(loadfile);
                        med1.SaveAs(copyFile);
                        string name = Path.GetFileName(fname);
                        string dir = newFolder + @"\" + name;
                        File.Move(fname, dir);
                    }
                }//END loop throught each file 
            }
            //Display all excel files
            string[] allFiles = Directory.GetFiles(@"C:\\Invoice\Excel Files");
            foreach (string ef in allFiles)
            {
                FileStream fs = File.Open(ef, FileMode.Open, FileAccess.Read);
                IExcelDataReader edr = ExcelReaderFactory.CreateOpenXmlReader(fs);
                result = edr.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });
                foreach (DataTable dt in result.Tables)
                {
                    Opts.Items.Add(dt.TableName);
                }
                edr.Close();
            }
        }

        private void Opts_SelectedIndexChanged(object sender, EventArgs e)
        {
            Display1.DataSource = result.Tables[Opts.SelectedIndex];
        }
    }
}
