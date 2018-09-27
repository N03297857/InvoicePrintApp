using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using ExcelDataReader;
using iTextSharp.text.pdf;
using iTextSharp.text;
namespace Trial_1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        protected List<DataSet> GenDS= new List<DataSet>(); // this variable holding the dataset for all excel files
        protected DataSet result = new DataSet(); // when you convert from excel file to dataset

        protected string addr1;
        protected int[] NMEDPreClean;
        protected int MedRow = 2; //keep tracking rows for Northern Medical Group
        protected List<string> NMedList = new List<string>(); // List of the the patient of and their statement of Northern Medical Group

        public NMGPDFGenerator createPDF;
        public NMGPatient newPatient;
        public CRSTCoverPage createCoverPage;
        public IDictionary<int, int> pageList;
        public IEnumerable<NMGPatient> patientData;
        public List<NMGPatientStatement> patientStatementList;
        public List<NMGPatient> patientList;

        public int amountOfPatients;
        public int amountOfPages;
        public IDictionary<int, int> coverPageValues;

        private void RDBtn_MouseClick(object sender, MouseEventArgs e)
        {
            string[] getFolder = Directory.GetDirectories(@"C:\\Invoice\Raw Data");
            foreach (string gF in getFolder)
            {
                switch (Path.GetFileName(gF))
                {
                    case "Northern Medical Group":
                          string[] getAllRawData = Directory.GetFiles(@"C:\\Invoice\Raw Data\" + Path.GetFileName(gF));
                          foreach(string gARD in getAllRawData)
                          {
                                NorthernMedicalGroupTXT(gARD);
                          }
                    break;
                }
            }
            //Display all excel files
            string[] allFiles = Directory.GetFiles(@"C:\\Invoice\Excel Files");
            foreach (string ef in allFiles)
            {
                FileStream fs = File.Open(ef, FileMode.Open, FileAccess.Read);
                IExcelDataReader edr = ExcelReaderFactory.CreateOpenXmlReader(fs);
                result = edr.AsDataSet();
                foreach (DataTable dt in result.Tables)
                {
                    Opts.Items.Add(dt.TableName);
                }
                edr.Close();
                GenDS.Add(result);
            }
            Opts.SelectedIndex = 0;
            Opts_SelectedIndexChanged(null, null);
            Display1.DataSource = GenDS[0].Tables[0];
        }

        private void Opts_SelectedIndexChanged(object sender, EventArgs e)
        {
            Display1.DataSource = GenDS[Opts.SelectedIndex].Tables[0];
        }

        private void ExcBtn_MouseClick(object sender, MouseEventArgs e)
        {
            string[] cleanFiles = Directory.GetFiles(@"C:\\Invoice\Clean Data");
            foreach (string cL in cleanFiles)
            {
                string fileName = Path.GetFileName(cL);
                string[] fileNameSplit = fileName.Split('_');
                switch (fileNameSplit[0])
                { 
                    case "NMED01":
                        NorthernMedicalGroupPDF(cL);
                    break;
                }
            }
        }

        private void NorthernMedicalGroupTXT(string textFile)
        {
            DateTime today = DateTime.Today;
            string date = today.ToString("dd-MM-yyyy");
            string[] dateSplit = date.Split('-');
            string[] lines = System.IO.File.ReadAllLines(textFile);
            NMedList = System.IO.File.ReadAllLines(textFile).ToList();
            string location = dateSplit[2] + @"\" + dateSplit[1] + @"\" + date;
            string newFolder = @"C:\Invoice\Clients\Northern Medical Group\" + location;
            string ExcelFolder = @"C:\Invoice\Excel Files";
            //create a folder based on current year, month, and date
            Directory.CreateDirectory(newFolder);
            //create a text file which have all the patients 
            string comText = newFolder + @"\NMED01_" + dateSplit[2] + "_" + dateSplit[1] + "_" + date + ".txt";
            //create the excel file of all the patients without the code 
            string ExcFile = newFolder + @"\NMED01_" + dateSplit[2] + "_" + dateSplit[1] + "_" + date + ".xlsx";
            //create the copy excel file of the all the patient so user can add bar code and tray number in 
            string ExcFile2 = ExcelFolder + @"\NMED01_" + dateSplit[2] + "_" + dateSplit[1] + "_" + date + ".xlsx";
            File.AppendAllLines(comText, NMedList);// keep adding the the text to the combine text file 
            //the excel file is not exist create one 
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
            FileInfo copyFile = new FileInfo(ExcFile2); // load copy file 
            ExcelPackage med1 = new ExcelPackage(loadfile);
            var worksheet1 = med1.Workbook.Worksheets["Northern Medical Group"];
            for (int i = 1; i < lines.Length; i++)
            {// read in each line
                if (lines[i] == "ecwPtStatement") // indicator for client format  
                {
                    var j = i + 1;
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
                            new string[]{info[4].Trim('"'),
                                        guid, info[0].Trim('"'),
                                        info[1].Trim('"'), info[2].Trim('"'),
                                        info[9].Trim('"'), info[10].Trim('"'),
                                        info[11].Trim('"'), info[12].Trim('"'),
                                        info[13].Trim('"')
                            }
                        };
                        worksheet1.Cells[MedRow, 1].LoadFromArrays(data);
                        MedRow++;
                    }
                }
            }
            med1.SaveAs(loadfile);
            med1.SaveAs(copyFile);
            string name = Path.GetFileName(textFile);
            string dir = newFolder + @"\" + name;
            //File.Move(textFile, dir);
        }

        private void NorthernMedicalGroupPDF(string cleanExcelFiles) 
        {
            string fileName = Path.GetFileName(cleanExcelFiles);
            string[] getFileName = fileName.Split('.');
            string[] fileNameSplit = getFileName[0].Split('_');
            string Direction = fileNameSplit[1] + @"\" + fileNameSplit[2] + @"\" + fileNameSplit[3];
            string FileName = fileNameSplit[1] + "_" + fileNameSplit[2] + "_" + fileNameSplit[3];
            string CompareTextFile = @"C:\Invoice\Clients\Northern Medical Group\" + Direction + @"\" + "NMED01_" + FileName + ".txt";
            string[] lines = System.IO.File.ReadAllLines(CompareTextFile);
            DataSet MEDDataSet = new DataSet();
            FileStream fs = File.Open(cleanExcelFiles, FileMode.Open, FileAccess.Read);
            IExcelDataReader MEDClean = ExcelReaderFactory.CreateOpenXmlReader(fs);
            MEDDataSet = MEDClean.AsDataSet();
            DataTable dt = MEDDataSet.Tables[0];
            //int count2 = 0;
            string resources = @"C:\Invoice\PDF Tools";
            string pdfFile = @"C:\Invoice\Clients\Northern Medical Group\" + Direction + @"\" + "NMED01_" + FileName + ".pdf";
            string coverPdfFile = @"C:\Invoice\Clients\Northern Medical Group\" + Direction + @"\";
            pageList = new Dictionary<int, int>();
            createPDF = new NMGPDFGenerator(resources);
            createCoverPage = new CRSTCoverPage(resources);
            patientList = new List<NMGPatient>(); //Stores different Patient
            patientStatementList = new List<NMGPatientStatement>(); //Stores each line of Patient Statement
            foreach (DataRow dr in dt.Rows)
            {
                for (int i = 1; i < lines.Length; i++)
                {// read in each line
                    if (lines[i] == "ecwPtStatement") // indicator for client format  
                    {
                        var j = i + 1;
                        if (j < lines.Length)
                        {
                            if(lines[j] == "ecwPtStatement,1.0") { j = j + 2; }
                            string[] info = lines[j].Split(',');
                            if (info[4].Trim('"') == dr["Column0"].ToString())
                            {
                                List<string> infoPatientList = new List<string>(info);
                                int curr = 0;
                                while (curr < infoPatientList.Count)
                                {
                                    int qCount = checkForOneQuote(infoPatientList[curr]);
                                    if (qCount < 2)
                                    {
                                        int nextWord = curr + 1;
                                        while (qCount < 2)
                                        {
                                            qCount = checkForOneQuote(infoPatientList[nextWord]);
                                            if (qCount >= 2) { break; }
                                            infoPatientList[curr] += infoPatientList[nextWord];
                                            infoPatientList.RemoveAt(nextWord);
                                        }
                                    }
                                    curr++;
                                }
                                //start
                                newPatient = new NMGPatient(info[4]);
                                string guid1 = System.Guid.NewGuid().ToString().Replace("-", "").ToUpper();
                                char charToTrim1 = '"';
                                var data1 = new List<string[]>()
                                    {
                                            new string[]{infoPatientList[4], guid1, infoPatientList[0], infoPatientList[1], infoPatientList[2], infoPatientList[3], infoPatientList[5],
                                            infoPatientList[6], infoPatientList[7], infoPatientList[8], infoPatientList[9], infoPatientList[10], infoPatientList[11], infoPatientList[12],
                                            infoPatientList[13], infoPatientList[14], infoPatientList[15], infoPatientList[16], infoPatientList[17], infoPatientList[18], infoPatientList[19],
                                            infoPatientList[20], infoPatientList[21], infoPatientList[22], infoPatientList[23], infoPatientList[24], infoPatientList[25], infoPatientList[26],
                                            infoPatientList[27], infoPatientList[28]}
                                    };
                                newPatient.PatientFirstName = infoPatientList[0].Trim(charToTrim1);
                                newPatient.PatientMiddleName = infoPatientList[1].Trim(charToTrim1);
                                newPatient.PatientLastName = infoPatientList[2].Trim(charToTrim1);
                                newPatient.PaymentDue = infoPatientList[5].Trim(charToTrim1);
                                int AcNo;
                                if (Int32.TryParse(infoPatientList[4].Trim(charToTrim1), out AcNo))
                                {
                                    newPatient.AccountNo = AcNo;
                                }
                                DateTime billDate;
                                if (DateTime.TryParse(infoPatientList[3].Trim(charToTrim1), out billDate))
                                {
                                    newPatient.BillDate = billDate;
                                }
                                newPatient.MailFirstName = infoPatientList[6].Trim(charToTrim1);
                                newPatient.MailMiddleName = infoPatientList[7].Trim(charToTrim1);
                                newPatient.MailLastName = infoPatientList[8].Trim(charToTrim1);
                                newPatient.MailAddressLine1 = infoPatientList[9].Trim(charToTrim1);
                                newPatient.MailAddressLine2 = infoPatientList[10].Trim(charToTrim1);
                                newPatient.MailCity = infoPatientList[11].Trim(charToTrim1);
                                newPatient.MailState = infoPatientList[12].Trim(charToTrim1);
                                newPatient.MailZip = infoPatientList[13].Trim(charToTrim1);
                                newPatient.RenderedName = infoPatientList[14].Trim(charToTrim1);
                                newPatient.RenderedAddressLine1 = infoPatientList[15].Trim(charToTrim1);
                                newPatient.RenderedAddressLine2 = infoPatientList[16].Trim(charToTrim1);
                                newPatient.RenderedCity = infoPatientList[17].Trim(charToTrim1);
                                newPatient.RenderedState = infoPatientList[18].Trim(charToTrim1);
                                newPatient.RenderedZip = infoPatientList[19].Trim(charToTrim1);
                                newPatient.PayableTo = infoPatientList[20].Trim(charToTrim1);
                                newPatient.Unknowing1 = infoPatientList[21].Trim(charToTrim1);
                                newPatient.Unknowing2 = infoPatientList[22].Trim(charToTrim1);
                                newPatient.AgingCurrent = infoPatientList[23].Trim(charToTrim1);
                                newPatient.Aging31_60 = infoPatientList[24].Trim(charToTrim1);
                                newPatient.Aging61_90 = infoPatientList[25].Trim(charToTrim1);
                                newPatient.Aging91_120 = infoPatientList[26].Trim(charToTrim1);
                                newPatient.Aging120 = infoPatientList[27].Trim(charToTrim1);
                                newPatient.InquireyPhone = infoPatientList[28].Trim(charToTrim1);
                                newPatient.IMBarcode = dr["Column10"].ToString();
                                patientList.Add(newPatient);
                                var h = j + 1;
                                if (lines.Length > h)
                                {
                                    while (lines[h] != "ecwPtStatement") //Iterates through each line for statement until it reaches next patient
                                    {
                                        if (h < lines.Length)
                                        {
                                            string[] infoPatientStatement = lines[h].Split(',');
                                            List<string> infoPatientStatementList = new List<string>(infoPatientStatement);
                                            int curr1 = 0;
                                            while (curr1 < infoPatientStatementList.Count)
                                            {
                                                int qCount1 = checkForOneQuote(infoPatientStatementList[curr1]);
                                                if (qCount1 < 2)
                                                {
                                                    int nextWord1 = curr1 + 1;
                                                    while (qCount1 < 2)
                                                    {
                                                        qCount1 = checkForOneQuote(infoPatientStatementList[nextWord1]);
                                                        if (qCount1 >= 2) { break; }
                                                        infoPatientStatementList[curr1] += infoPatientStatementList[nextWord1];
                                                        infoPatientStatementList.RemoveAt(nextWord1);
                                                    }
                                                }
                                                curr1++;
                                            }
                                            string guid2 = System.Guid.NewGuid().ToString().Replace("-", "").ToUpper();
                                            char charToTrim2 = '"';
                                            var data2 = new List<string[]>()
                                            {
                                                new string[]{infoPatientStatementList[0], infoPatientStatementList[1], infoPatientStatementList[2],
                                                infoPatientStatementList[3], infoPatientStatementList[4], infoPatientStatementList[5], infoPatientStatementList[6],
                                                infoPatientStatementList[7]}
                                            };
                                            NMGPatientStatement newPatientStatement = new NMGPatientStatement();
                                            newPatientStatement.AccountNo = infoPatientStatementList[0].Trim(charToTrim2);
                                            newPatientStatement.ClaimNo = infoPatientStatementList[1].Trim(charToTrim2);
                                            DateTime ViDate;
                                            DateTime AcDate;
                                            if (DateTime.TryParse(infoPatientStatementList[2].Trim(charToTrim2), out ViDate))
                                            {
                                                newPatientStatement.VisitDate = ViDate;
                                            }
                                            if (DateTime.TryParse(infoPatientStatementList[3].Trim(charToTrim2), out AcDate))
                                            {
                                                newPatientStatement.ActivityDate = AcDate;
                                            }
                                            newPatientStatement.SetDescription(infoPatientStatementList[4].Trim(charToTrim2));
                                            newPatientStatement.Charges = infoPatientStatementList[5].Trim(charToTrim2);
                                            newPatientStatement.Payments = infoPatientStatementList[6].Trim(charToTrim2);
                                            newPatientStatement.Balance = infoPatientStatementList[7].Trim(charToTrim2);
                                            patientStatementList.Add(newPatientStatement);
                                            h++;
                                        }
                                    }
                                    newPatient.SetStatement(patientStatementList);
                                }
                                int patientStatementListSize = patientStatementList.Count;
                                patientStatementList.RemoveRange(0, patientStatementListSize);
                            }
                        }              
                    }
                }
            }
            //Creates PDF for all patients into one file
            createPDF.GeneratorPDF(patientList, pdfFile);
            //Finds total amount of pages and patients for PDF file
            PdfReader pdfRead = new PdfReader(pdfFile);
            amountOfPages = pdfRead.NumberOfPages;
            amountOfPatients = patientList.Count;
            //Splits file to load into CRSTCoverPage
            string backSlash = "\"";
            string[] fileLines = Regex.Split(coverPdfFile, backSlash);
            //Grabs total pages for individual patients for CRSTCoverPage 
            pageList.Add(amountOfPages, amountOfPatients);
            //Prints cover page for Northern Medical Group
            createCoverPage.PrintCoverPage(pageList, amountOfPatients, amountOfPages, fileLines, coverPdfFile);
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
    }
}
