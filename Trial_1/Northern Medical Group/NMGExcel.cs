using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using ExcelDataReader;

namespace Trial_1.Northern_Medical_Group
{
    class NMGExcel
    {
        protected int generalCount = 0;
        protected int page; protected int sta;
        protected List<DataSet> GenDS = new List<DataSet>(); // this variable holding the dataset for all excel files
        protected DataSet result = new DataSet(); // when you convert from excel file to dataset
        public Form1 form = new Form1();

        protected string MedExcFile;
        protected int MedRow = 2; //keep tracking rows for Northern Medical Group
        protected int MedRow2 = 2; //keep tracking rows for Northern Medical Group
        protected List<string> NMedList = new List<string>(); // List of the the patient of and their statement of Northern Medical Group

        public string ToDateFolder;
        public string RawDataFolder;
        public string ExcelDataFolder;
        public string CleanDataFolder;
        public string MailDatFolder;
        public string ReturnFolder;

        public void NorthernMedicalGroupTXT(string textFile)
        {
            CreateFolder("Northern Medical Group");
            DateTime today = DateTime.Today;
            string date = today.ToString("MMMM-dd-yyyy-MM");
            string[] lines = System.IO.File.ReadAllLines(textFile);
            NMedList = System.IO.File.ReadAllLines(textFile).ToList();
            string ExcelFolder = @"L:\__Invoice\Excel Files\Northern Medical Group";
            // create a text file which have all the patients 
            string comText = CleanDataFolder + @"\NMG_" + date + ".txt";
            //create the excel file of all the patients without the code 
            string ExcFile = ExcelDataFolder + @"\NMG_" + date + ".xlsx";
            // create the copy excel file of the all the patient so user can add bar code and tray number in 
            MedExcFile = ExcelFolder + @"\NMG_" + date + ".xlsx";
            // create the excel file for all patients who have more than 3 pages 
            string ExcFile3 = ExcelDataFolder + @"\NMG_" + date + "_Extra.xlsx";
            File.AppendAllLines(comText, NMedList);// keep adding the the text to the combine text file 
            //the excel file is not exist create one 
            if (!System.IO.File.Exists(ExcFile))
            {
                ExcelPackage med = new ExcelPackage();
                med.Workbook.Worksheets.Add("Northern Medical Group");//creat work sheet
                var headerRow = new List<string[]>()//add header
                            {
                                new string[] { "ID","ACCOUNT", "FIRST NAME", "MIDDLE NAME", "LAST NAME", "ADDRESSLINE1", "ADDRESSLINE2", "CITY", "STATE", "ZIPCODE", "PAGES" }
                            };
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                var worksheet = med.Workbook.Worksheets["Northern Medical Group"]; // load worksheet based on name
                worksheet.Cells[headerRange].LoadFromArrays(headerRow); //load data to worksheet 
                FileInfo excelFile = new FileInfo(ExcFile);
                FileInfo excelFile2 = new FileInfo(MedExcFile);
                FileInfo excelFile3 = new FileInfo(ExcFile3);
                med.SaveAs(excelFile); // save file
                med.SaveAs(excelFile2);
                med.SaveAs(excelFile3);
            }
            FileInfo loadfile = new FileInfo(ExcFile);// load excel file
            FileInfo copyFile = new FileInfo(MedExcFile); // load copy file 
            FileInfo ExtraFile = new FileInfo(ExcFile3); // load copy file 
            ExcelPackage med1 = new ExcelPackage(loadfile);
            ExcelPackage med2 = new ExcelPackage(ExtraFile);
            var worksheet2 = med2.Workbook.Worksheets["Northern Medical Group"];
            var worksheet1 = med1.Workbook.Worksheets["Northern Medical Group"];
            for (int i = 1; i < lines.Length; i++)
            {// read in each line
                if (lines[i] == "ecwPtStatement") // indicator for client format  
                {
                    page = 1; sta = 1;
                    var j = i + 1;
                    if (lines[j] == "ecwPtStatement,1.0") { j = j + 2; }
                    var c = j + 1;
                    while (lines[c] != "ecwPtStatement" || lines[c] != "ecwPtStatement,1.0")
                    {
                        sta++;
                        if (sta > 30)
                        {
                            page++;
                            sta = 1;
                        }
                        c++;
                        if (c < lines.Length)
                        {
                            if (lines[c] == "ecwPtStatement" || lines[c] == "ecwPtStatement,1.0") { break; }
                        }
                        else { break; }
                    }
                    if (j < lines.Length)
                    {
                        string[] infoArr = lines[j].Split(',');
                        string guid = System.Guid.NewGuid().ToString().Replace("-", "").ToUpper();
                        List<string> info = new List<string>(infoArr);
                        string checkTrim = info[14].Trim('"');
                        if (checkTrim != "Northern Medical Group")
                        {
                            int curr = 0;
                            while (curr < info.Count)
                            {
                                int qCount = checkForOneQuote(info[curr]);
                                if (qCount < 2)
                                {
                                    int nextWord = curr + 1;
                                    while (qCount < 2)
                                    {
                                        qCount = checkForOneQuote(info[nextWord]);
                                        if (qCount >= 2) { break; }
                                        info[curr] += info[nextWord];
                                        info.RemoveAt(nextWord);
                                    }
                                }
                                curr++;
                                checkTrim = info[curr].Trim('"');
                                if (checkTrim == "Northern Medical Group")
                                {
                                    break;
                                }
                            }
                        }
                        var data = new List<string[]>()
                        {
                            new string[]{guid, info[4].Trim('"'),
                                        info[0].Trim('"'), info[1].Trim('"'),
                                        info[2].Trim('"'), info[9].Trim('"'),
                                        info[10].Trim('"'), info[11].Trim('"'),
                                        info[12].Trim('"'), info[13].Trim('"'), page.ToString()
                            }
                        };
                        if (page > 3)
                        {
                            worksheet2.Cells[MedRow2, 1].LoadFromArrays(data);
                            MedRow2++;
                        }
                        else
                        {
                            worksheet1.Cells[MedRow, 1].LoadFromArrays(data);
                            MedRow++;
                        }
                    }
                }
            }
            med1.SaveAs(loadfile);
            med1.SaveAs(copyFile);
            med2.SaveAs(ExtraFile);
            string name = Path.GetFileName(textFile);
            string dir = RawDataFolder + @"\" + name;
            File.Move(textFile, dir);
        }

        private int checkForOneQuote(string checkString)
        {
            int count = 0;
            foreach (var q in checkString)
            {
                if (q == '"')
                {
                    count++;
                }
            }
            return count;
        }

        public void CreateFolder(string ClientName)
        {
            DateTime today = DateTime.Today;
            string date = today.ToString("MMMM-dd-yyyy-MM");
            string[] dateSplit = date.Split('-');
            string FolderDate = dateSplit[0] + "-" + dateSplit[1] + "-" + dateSplit[2];
            string location = dateSplit[2] + @"\" + dateSplit[0] + @"\" + FolderDate + " Invoices";
            ToDateFolder = @"L:\__Invoice\Clients\" + ClientName + @"\" + location;
            Directory.CreateDirectory(ToDateFolder);
            RawDataFolder = ToDateFolder + @"\RawData";
            Directory.CreateDirectory(RawDataFolder);
            ExcelDataFolder = ToDateFolder + @"\ExcelData";
            Directory.CreateDirectory(ExcelDataFolder);
            CleanDataFolder = ToDateFolder + @"\CleanData";
            Directory.CreateDirectory(CleanDataFolder);
            MailDatFolder = ToDateFolder + @"\Mail.Dat";
            Directory.CreateDirectory(MailDatFolder);
            ReturnFolder = ToDateFolder + @"\Returns";
            Directory.CreateDirectory(ReturnFolder);
        }
    }
}
