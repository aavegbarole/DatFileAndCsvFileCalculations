namespace MergeDL_DAT_V2P_Data
{
    using LINQtoCSV;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Windows.Forms;

    /// Main form for the Application
    /// Defines the <see cref="Form1" />
    /// 
    public partial class Form1 : Form
    {
        /// Main form for the Application having 4 buttons and 3 Text box
        /// Initializes a new instance of the <see cref="Form1"/> class.
        /// 
        public Form1()
        {
            InitializeComponent();
            button1.Enabled = true;
            textBox1.Enabled = true;

            button2.Enabled = false;
            textBox2.Enabled = false;

            button3.Enabled = false;
            textBox3.Enabled = false;

            button4.Enabled = false;
        }

        /// <summary>
        /// The button1_Click
        /// This buttons is used to select Updated V2P File (CSV format)
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/></param>
        /// <param name="e">The e<see cref="EventArgs"/></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Select the Updated V2P File";
            fdlg.Filter = "CSV files (*.csv)|*.csv";
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fdlg.FileName;
                button2.Enabled = true;
                textBox2.Enabled = true;
            }
        }

        /// <summary>
        /// The button2_Click
        /// This buttons is used to select DAT File (GPS data)
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/></param>
        /// <param name="e">The e<see cref="EventArgs"/></param>
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Select the DAT File";
            fdlg.Filter = "DATA files (*.dat)|*.dat";
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = fdlg.FileName;
                button3.Enabled = true;
                textBox3.Enabled = true;
            }
        }

        /// <summary>
        /// The button3_Click
        /// /// This buttons is used to select the csv file having result data of Deep learning
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/></param>
        /// <param name="e">The e<see cref="EventArgs"/></param>
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Select the Deep Learning O/P CSV File";
            fdlg.Filter = "CSV files (*.csv)|*.csv";
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = fdlg.FileName;
                button4.Enabled = true;
            }
        }

        //*********** V2P Csv file list ************************************************
        /// <summary>
        /// ImageFileName stores the name of images (Ex. v2photo00002.jpg)
        /// </summary>
        internal List<String> ImageFileName = new List<string>();

        /// <summary>
        /// ShootDate stores the date when the video/image was taken (Ex. 2018/10/27)
        /// </summary>
        internal List<String> ShootDate = new List<string>();

        /// <summary>
        /// TimeStamp stores the Time when the video/image was taken (Ex. 15:04:06)
        /// TimeStamp doesnot stores the data it stores only TIME
        /// </summary>
        internal List<String> TimeStamp = new List<string>();

        /// <summary>
        /// Latitude stores latitude value at that point(image) 
        /// </summary>
        internal List<Decimal> Latitude = new List<Decimal>();

        /// <summary>
        /// Longitude stores longitude value at that point(image) 
        /// </summary>
        internal List<Decimal> Longitude = new List<Decimal>();

        //************ DAT file list****************************************************
        /// <summary>
        /// Speed stores speed of vehical at thet point(image)
        /// </summary>
        internal List<Decimal> Speed = new List<Decimal>();

        /// <summary>
        /// TimeStampDAT stores TimeStamp value from the DAT file (GPS data)
        /// </summary>
        internal List<String> TimeStampDAT = new List<String>();

        //*************** ML CSV File list *********************************************
        /// <summary>
        /// Stores the linear_crack values
        /// </summary>
        internal List<String> linear_crack = new List<String>();

        /// <summary>
        /// Stores the joint  values
        /// </summary>
        internal List<String> joint = new List<String>();

        /// <summary>
        /// Stores the manhole values
        /// </summary>
        internal List<String> manhole = new List<String>();

        /// <summary>
        /// Stores the alligator_crack  values
        /// </summary>
        internal List<String> alligator_crack = new List<String>();

        /// <summary>
        /// Stores the patch values
        /// </summary>
        internal List<String> patch = new List<String>();

        /// <summary>
        /// Stores the white_line values
        /// </summary>
        internal List<String> white_line = new List<String>();

        /// <summary>
        /// Stores the long_white_line values
        /// </summary>
        internal List<String> long_white_line = new List<String>();

        /// <summary>
        /// Stores the grating values
        /// </summary>
        internal List<String> grating = new List<String>();

        /// <summary>
        /// Stores the linear_crack_p values
        /// </summary>
        internal List<String> linear_crack_p = new List<String>();

        /// <summary>
        /// Stores the alligator_crack_p values
        /// </summary>
        internal List<String> alligator_crack_p = new List<String>();

        /// <summary>
        /// Stores the recess values
        /// </summary>
        internal List<String> recess = new List<String>();

        /// <summary>
        /// Stores the rutting values
        /// </summary>
        internal List<String> rutting = new List<String>();

        /// <summary>
        /// Stores the pothole values
        /// </summary>
        internal List<String> pothole = new List<String>();

        /// <summary>
        /// Stores the construction_j values
        /// </summary>
        internal List<String> construction_j = new List<String>();

        /// <summary>
        /// Stores the p_longitudinal values
        /// </summary>
        internal List<String> p_longitudinal = new List<String>();

        /// <summary>
        /// Stores the p_transverse values
        /// </summary>
        internal List<String> p_transverse = new List<String>();

        /// <summary>
        /// Stores the w_line_crack values
        /// </summary>
        internal List<String> w_line_crack = new List<String>();

        /// <summary>
        /// Stores the w_alligator_crack values
        /// </summary>
        internal List<String> w_alligator_crack = new List<String>();

        /// <summary>
        /// Stores the y_line_crack values
        /// </summary>
        internal List<String> y_line_crack = new List<String>();

        /// <summary>
        /// Stores the y_alligator_crack values
        /// </summary>
        internal List<String> y_alligator_crack = new List<String>();

        /// <summary>
        /// The button4_Click
        /// Starts the calculation and execution using the the aboue three selected files
        /// </summary>
        /// <param name="sender">The sender<see cref="object"/></param>
        /// <param name="e">The e<see cref="EventArgs"/></param>
        private void button4_Click(object sender, EventArgs e)
        {
            //*********** Start Reading V2p file**********************
            string pathOfV2PcsvFile = textBox1.Text;
            string pathOfDLcsvFile = textBox3.Text;

            CsvContext cc1 = new CsvContext();
            IEnumerable<CsvData> csvClassObject = cc1.Read<CsvData>(pathOfV2PcsvFile, CsvFileDescription);

            var tryQuery1 = from aData in csvClassObject
                            select new { aData };

            foreach (var Obj in tryQuery1)
            {
                ImageFileName.Add(Obj.aData.ImageFileName);
                ShootDate.Add((Obj.aData.ShootDate.Year + "/" + Obj.aData.ShootDate.Month + "/" + Obj.aData.ShootDate.Day).ToString());
                TimeStamp.Add((Obj.aData.TimeStamp.TimeOfDay.Hours + ":" + Obj.aData.TimeStamp.TimeOfDay.Minutes + ":" + Obj.aData.TimeStamp.TimeOfDay.Seconds).ToString());
                Latitude.Add(Obj.aData.Latitude);
                Longitude.Add(Obj.aData.Longitude);
            }

            //*********** END Reading V2P file**********************


            //*********** Start Reading DAT file**********************

            //** Covert DAT file to CSV file [START]
            string filePathWithName = textBox2.Text;
            string[] filePathArr = filePathWithName.Split(new char[] { '\\' });
            string fileName = filePathArr[(filePathArr.Length - 1)];
            string folderPath = filePathWithName.Remove(filePathWithName.Length - fileName.Length);

            System.IO.StreamReader reader = new System.IO.StreamReader(textBox2.Text);
            string strAllFile = reader.ReadToEnd().Replace("\r\n", "\n");
            string[] arrLines = strAllFile.Split(new char[] { '\n' });

            string datToCsvFileFullpathWithFileName = folderPath + "DatToCsvFile_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";

            System.IO.StreamWriter writer = new System.IO.StreamWriter(datToCsvFileFullpathWithFileName);
            for (int i = 2; i < arrLines.Length; i++)
            {
                writer.WriteLine(arrLines[i]);
            }
            writer.Close();
            //** Covert DAT file to CSV file [END]

            CsvContext cc2 = new CsvContext();
            IEnumerable<DatFileData> datClassObject = cc2.Read<DatFileData>(datToCsvFileFullpathWithFileName, datInputFileDescription);

            var tryQuery2 = from aData in datClassObject
                            select new { aData };

            foreach (var Obj in tryQuery2)
            {
                if (TimeStamp.Contains((Obj.aData.gpsTimeStamp.TimeOfDay.Hours + ":" + Obj.aData.gpsTimeStamp.TimeOfDay.Minutes + ":" + Obj.aData.gpsTimeStamp.TimeOfDay.Seconds).ToString()))
                {
                    if (TimeStampDAT.Count == 0)
                    {
                        Speed.Add(Obj.aData.speed);
                        TimeStampDAT.Add((Obj.aData.gpsTimeStamp.TimeOfDay.Hours + ":" + Obj.aData.gpsTimeStamp.TimeOfDay.Minutes + ":" + Obj.aData.gpsTimeStamp.TimeOfDay.Seconds).ToString());
                    }
                    else
                    {

                        if (!TimeStampDAT[TimeStampDAT.Count - 1].Equals(Obj.aData.gpsTimeStamp.TimeOfDay.Hours + ":" + Obj.aData.gpsTimeStamp.TimeOfDay.Minutes + ":" + Obj.aData.gpsTimeStamp.TimeOfDay.Seconds))
                        {
                            Speed.Add(Obj.aData.speed);
                            TimeStampDAT.Add((Obj.aData.gpsTimeStamp.TimeOfDay.Hours + ":" + Obj.aData.gpsTimeStamp.TimeOfDay.Minutes + ":" + Obj.aData.gpsTimeStamp.TimeOfDay.Seconds).ToString());
                        }
                    }
                }
            }

            loadCsvFile(pathOfDLcsvFile);
            int totalNumberOfrows = ImageFileName.Count;

            //*********** END Reading DAT file**********************

            //*********** Start Writing CSV Fille ******************
            var csv = new StringBuilder();

            for (var i = 0; i < totalNumberOfrows; i++)
            {
                var one = ImageFileName[i].ToString();
                var two = ShootDate[i].ToString();
                var three = TimeStamp[i].ToString();
                var four = Latitude[i].ToString();
                var five = Longitude[i].ToString();
                var six = Speed[i].ToString();
                var seven = linear_crack[i].ToString();
                var eight = joint[i].ToString();
                var nine = manhole[i].ToString();
                var ten = alligator_crack[i].ToString();
                var eleven = patch[i].ToString();
                var twelve = white_line[i].ToString();
                var thirteen = long_white_line[i].ToString();
                var fourtheen = grating[i].ToString();
                var fifteen = linear_crack_p[i].ToString();
                var sixteen = alligator_crack_p[i].ToString();
                var seventeen = recess[i].ToString();
                var eighteen = rutting[i].ToString();
                var ninteen = pothole[i].ToString();
                var twenty = construction_j[i].ToString();
                var twentyOne = p_longitudinal[i].ToString();
                var twentyTwo = p_transverse[i].ToString();
                var twentyThree = w_line_crack[i].ToString();
                var twentyFour = w_alligator_crack[i].ToString();
                var twentyFive = y_line_crack[i].ToString();
                var twentySix = y_alligator_crack[i].ToString();

                string[] array = new string[] { one, two, three, four, five, six, seven, eight, nine, ten, eleven, twelve, thirteen, fourtheen, fifteen, sixteen, seventeen, eighteen, ninteen, twenty, twentyOne, twentyTwo, twentyThree, twentyFour, twentyFive, twentySix };

                string refFormat = "{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}";

                string newLine = String.Format(refFormat, array);


                //var newLine = string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}", one, two, three, four, five, six, seven, eight, nine, ten, eleven, twelve, thirteen, fourtheen, fifteen, sixteen, seventeen, eighteen, twenty, twentyOne, twentyTwo, twentyThree, twentyFour, twentyFive, twentySix, twentySeven);
                csv.AppendLine(newLine);
            }
            string ResultFileNameCSV = folderPath + "V2P_DAT_ML_CSV" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv";
            File.WriteAllText(ResultFileNameCSV, csv.ToString());

            //*********** END Writing CSV Fille ******************

            //*********** Start Writing Coloured/Detailed Excel Fille ******************
            // Load Excel application
            Microsoft.Office.Interop.Excel.Application excelWrite = new Microsoft.Office.Interop.Excel.Application();
            excelWrite.DisplayAlerts = false;
            // Create empty workbook
            Microsoft.Office.Interop.Excel.Workbook wb = excelWrite.Workbooks.Add();
            //excelWrite.DefaultSaveFormat

            // Create Worksheet from active sheet
            Microsoft.Office.Interop.Excel.Worksheet workSheet = wb.ActiveSheet;
            workSheet.Name = "Detailed_Result";

            //**** Merging of header cells
            workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1, 5]].Merge();
            workSheet.Range[workSheet.Cells[1, 7], workSheet.Cells[1, 26]].Merge();

            //************** Range of V2P data and adding color and formatting to it
            int startColumn1 = 1;
            int startRow1 = 1;
            Microsoft.Office.Interop.Excel.Range startCell1 = workSheet.Cells[startRow1, startColumn1];

            int endColumn1 = 5;
            int endRow1 = totalNumberOfrows + 2;
            Microsoft.Office.Interop.Excel.Range endCell1 = workSheet.Cells[endRow1, endColumn1];

            Microsoft.Office.Interop.Excel.Range V2PcsvRange = workSheet.Range[startCell1, endCell1];
            V2PcsvRange.Columns.AutoFit();
            V2PcsvRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.BlanchedAlmond);

            V2PcsvRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            V2PcsvRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            V2PcsvRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            V2PcsvRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);

            //************** Range of DAT data and adding color and formatting to it
            int startColumn2 = 6;
            int startRow2 = 1;
            Microsoft.Office.Interop.Excel.Range startCell2 = workSheet.Cells[startRow2, startColumn2];

            int endColumn2 = 6;
            int endRow2 = totalNumberOfrows + 2;
            Microsoft.Office.Interop.Excel.Range endCell2 = workSheet.Cells[endRow2, endColumn2];

            Microsoft.Office.Interop.Excel.Range DatRange = workSheet.Range[startCell2, endCell2];
            DatRange.Columns.AutoFit();
            DatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);

            DatRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            DatRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            DatRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            DatRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);

            //************** Range of Deep Learning data and adding color and formatting to it
            int startColumn3 = 7;
            int startRow3 = 1;
            Microsoft.Office.Interop.Excel.Range startCell3 = workSheet.Cells[startRow3, startColumn3];

            int endColumn3 = 26;
            int endRow3 = totalNumberOfrows + 2;
            Microsoft.Office.Interop.Excel.Range endCell3 = workSheet.Cells[endRow3, endColumn3];

            Microsoft.Office.Interop.Excel.Range DLRange = workSheet.Range[startCell3, endCell3];
            DLRange.Columns.AutoFit();
            DLRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LemonChiffon);

            DLRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            DLRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            DLRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            DLRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);

            //************** Range of Longitude and latitude in DMS format data and adding color and formatting to it
            int startColumnDMS = 27;
            int startRowDMS = 2;
            Microsoft.Office.Interop.Excel.Range startCellDMS = workSheet.Cells[startRowDMS, startColumnDMS];

            int endColumnDMS = 28;
            int endRowDMS = totalNumberOfrows + 2;
            Microsoft.Office.Interop.Excel.Range endCellDMS = workSheet.Cells[endRowDMS, endColumnDMS];

            Microsoft.Office.Interop.Excel.Range DMSDataRange = workSheet.Range[startCellDMS, endCellDMS];
            DMSDataRange.Columns.AutoFit();
            DMSDataRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MintCream);

            DMSDataRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            DMSDataRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            DMSDataRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            DMSDataRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);

            //**** Writing heading directly in file
            workSheet.Cells[1, "A"] = "V2P File Data";
            workSheet.Cells[1, "A"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            workSheet.Cells[1, "F"] = "DAT File Data";
            workSheet.Cells[1, "F"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            workSheet.Cells[1, "G"] = "DEEP LEARNING Result Data";
            workSheet.Cells[1, "G"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            workSheet.Cells[2, "A"] = "Image File Name";
            workSheet.Cells[2, "B"] = "Date";
            workSheet.Cells[2, "C"] = "Time Stamp";
            workSheet.Cells[2, "D"] = "Latitude";
            workSheet.Cells[2, "E"] = "Longitude";

            workSheet.Cells[2, "F"] = "Speed";

            workSheet.Cells[2, "AA"] = "Latitude (DMS)";
            workSheet.Cells[2, "AB"] = "Longitude (DMS)";

            //*********** Adding hyperlinks to Deep learning data headings (0, 1, 2, ......)
            workSheet.Cells[2, "G"] = "0";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink0 = workSheet.get_Range("G2", Type.Missing);
            string hyperlinkTargetAddress0 = "Deep_Learning_Column_Names!B1";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink0, string.Empty, hyperlinkTargetAddress0, "0: linear crack  線状ひび割れ");

            workSheet.Cells[2, "H"] = "1";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink1 = workSheet.get_Range("H2", Type.Missing);
            string hyperlinkTargetAddress1 = "Deep_Learning_Column_Names!B2";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink1, string.Empty, hyperlinkTargetAddress1, "1: joint  ジョイント");

            workSheet.Cells[2, "I"] = "2";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink2 = workSheet.get_Range("I2", Type.Missing);
            string hyperlinkTargetAddress2 = "Deep_Learning_Column_Names!B3";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink2, string.Empty, hyperlinkTargetAddress2, "2: manhole  マンホール");

            workSheet.Cells[2, "J"] = "3";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink3 = workSheet.get_Range("J2", Type.Missing);
            string hyperlinkTargetAddress3 = "Deep_Learning_Column_Names!B4";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink3, string.Empty, hyperlinkTargetAddress3, "3: alligator crack  亀甲状ひび割れ");

            workSheet.Cells[2, "K"] = "4";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink4 = workSheet.get_Range("K2", Type.Missing);
            string hyperlinkTargetAddress4 = "Deep_Learning_Column_Names!B5";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink4, string.Empty, hyperlinkTargetAddress4, "4: patch  パッチング");

            workSheet.Cells[2, "L"] = "5";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink5 = workSheet.get_Range("L2", Type.Missing);
            string hyperlinkTargetAddress5 = "Deep_Learning_Column_Names!B6";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink5, string.Empty, hyperlinkTargetAddress5, "5: white line  短い白線");

            workSheet.Cells[2, "M"] = "6";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink6 = workSheet.get_Range("M2", Type.Missing);
            string hyperlinkTargetAddress6 = "Deep_Learning_Column_Names!B7";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink6, string.Empty, hyperlinkTargetAddress6, "6: long white_line  長い白線");

            workSheet.Cells[2, "N"] = "7";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink7 = workSheet.get_Range("N2", Type.Missing);
            string hyperlinkTargetAddress7 = "Deep_Learning_Column_Names!B8";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink7, string.Empty, hyperlinkTargetAddress7, "7: grating  グレーチング");

            workSheet.Cells[2, "O"] = "8";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink8 = workSheet.get_Range("O2", Type.Missing);
            string hyperlinkTargetAddress8 = "Deep_Learning_Column_Names!B9";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink8, string.Empty, hyperlinkTargetAddress8, "8: linear crack_p  線状ひび割れ補修跡");

            workSheet.Cells[2, "P"] = "9";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink9 = workSheet.get_Range("P2", Type.Missing);
            string hyperlinkTargetAddress9 = "Deep_Learning_Column_Names!B10";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink9, string.Empty, hyperlinkTargetAddress9, "9: alligator crack_p 亀甲状ひび割れ補修跡");

            workSheet.Cells[2, "Q"] = "10";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink10 = workSheet.get_Range("Q2", Type.Missing);
            string hyperlinkTargetAddress10 = "Deep_Learning_Column_Names!B11";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink10, string.Empty, hyperlinkTargetAddress10, "10: recess  剥離・くぼみ");

            workSheet.Cells[2, "R"] = "11";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink11 = workSheet.get_Range("R2", Type.Missing);
            string hyperlinkTargetAddress11 = "Deep_Learning_Column_Names!B12";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink11, string.Empty, hyperlinkTargetAddress11, "11: rutting  わだち掘れ");

            workSheet.Cells[2, "S"] = "12";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink12 = workSheet.get_Range("S2", Type.Missing);
            string hyperlinkTargetAddress12 = "Deep_Learning_Column_Names!B13";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink12, string.Empty, hyperlinkTargetAddress12, "12: pothole  ポットホール");

            workSheet.Cells[2, "T"] = "13";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink13 = workSheet.get_Range("T2", Type.Missing);
            string hyperlinkTargetAddress13 = "Deep_Learning_Column_Names!B14";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink13, string.Empty, hyperlinkTargetAddress13, "13: construction_j  施工上の継ぎ目");

            workSheet.Cells[2, "U"] = "14";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink14 = workSheet.get_Range("U2", Type.Missing);
            string hyperlinkTargetAddress14 = "Deep_Learning_Column_Names!B15";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink14, string.Empty, hyperlinkTargetAddress14, "14: p_longitudinal  パッチング縦");

            workSheet.Cells[2, "V"] = "15";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink15 = workSheet.get_Range("V2", Type.Missing);
            string hyperlinkTargetAddress15 = "Deep_Learning_Column_Names!B16";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink15, string.Empty, hyperlinkTargetAddress15, "15: p_transverse  パッチング横");

            workSheet.Cells[2, "W"] = "16";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink16 = workSheet.get_Range("W2", Type.Missing);
            string hyperlinkTargetAddress16 = "Deep_Learning_Column_Names!B17";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink16, string.Empty, hyperlinkTargetAddress16, "16: w_line_crack  白線上の線状ひび割れ");

            workSheet.Cells[2, "X"] = "17";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink17 = workSheet.get_Range("X2", Type.Missing);
            string hyperlinkTargetAddress17 = "Deep_Learning_Column_Names!B18";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink17, string.Empty, hyperlinkTargetAddress17, "17: w_alligator_crack  白線上の亀甲状ひび割れ");

            workSheet.Cells[2, "Y"] = "18";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink18 = workSheet.get_Range("Y2", Type.Missing);
            string hyperlinkTargetAddress18 = "Deep_Learning_Column_Names!B19";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink18, string.Empty, hyperlinkTargetAddress18, "18: y_line_crack  黄線上の線状ひび割れ");

            workSheet.Cells[2, "Z"] = "19";
            Microsoft.Office.Interop.Excel.Range rangeToHoldHyperlink19 = workSheet.get_Range("Z2", Type.Missing);
            string hyperlinkTargetAddress19 = "Deep_Learning_Column_Names!B20";
            workSheet.Hyperlinks.Add(rangeToHoldHyperlink19, string.Empty, hyperlinkTargetAddress19, "19: y_alligator_crack  黄線上の亀甲状ひび割れ");

            //************ Code foe writting the data in file from the lists where data was saved while reading Using LINQ
            for (var n = 0; n < totalNumberOfrows; n++)
            {
                workSheet.Cells[n + 3, "A"] = ImageFileName[n].ToString();
                workSheet.Cells[n + 3, "B"] = ShootDate[n].ToString();
                workSheet.Cells[n + 3, "C"] = TimeStamp[n].ToString();
                workSheet.Cells[n + 3, "D"] = Latitude[n].ToString();
                workSheet.Cells[n + 3, "E"] = Longitude[n].ToString();
                workSheet.Cells[n + 3, "F"] = Speed[n].ToString();

                //************ Formula for latitude and longitude in DMS format
                workSheet.Cells[n + 3, "AA"].value2 = "=CONCATENATE(TEXT(ROUNDDOWN(ABS(D" + Convert.ToDecimal(n + 3) + "),0),\"00\"),\"° \",TEXT(ROUNDDOWN(ABS((D" + Convert.ToDecimal(n + 3) + "-ROUNDDOWN(D" + Convert.ToDecimal(n + 3) + ",0))*60),0),\"00\"),\"' \",TEXT(TRUNC((ABS((D" + Convert.ToDecimal(n + 3) + "-ROUNDDOWN(D" + Convert.ToDecimal(n + 3) + ",0))*60)-ROUNDDOWN(ABS((D" + Convert.ToDecimal(n + 3) + "-ROUNDDOWN(D" + Convert.ToDecimal(n + 3) + ",0))*60),0))*60,2),\"00.00\"),\"\"\"\",IF(D" + Convert.ToDecimal(n + 3) + "<0,\" S\",\" N\"))";
                workSheet.Cells[n + 3, "AB"] = "=CONCATENATE(TEXT(ROUNDDOWN(ABS(E" + Convert.ToDecimal(n + 3) + "),0),\"000\"),\"° \",TEXT(ROUNDDOWN(ABS((E" + Convert.ToDecimal(n + 3) + "-ROUNDDOWN(E" + Convert.ToDecimal(n + 3) + ",0))*60),0),\"00\"),\"' \",TEXT(TRUNC((ABS((E" + Convert.ToDecimal(n + 3) + "-ROUNDDOWN(E" + Convert.ToDecimal(n + 3) + ",0))*60)-ROUNDDOWN(ABS((E" + Convert.ToDecimal(n + 3) + "-ROUNDDOWN(E" + Convert.ToDecimal(n + 3) + ",0))*60),0))*60,2),\"00.00\"),\"\"\"\",IF(E" + Convert.ToDecimal(n + 3) + "<0,\" W\",\" E\"))";

                //****** Condition to check if the data in not zero and mark it with another color (DL data)
                if (Convert.ToDecimal(linear_crack[n]) != 0)
                {
                    workSheet.Cells[n + 3, "G"] = linear_crack[n].ToString();
                    workSheet.Cells[n + 3, "G"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "G"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "G"] = linear_crack[n].ToString();
                }

                if (Convert.ToDecimal(joint[n]) != 0)
                {
                    workSheet.Cells[n + 3, "H"] = joint[n].ToString();
                    workSheet.Cells[n + 3, "H"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "H"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "H"] = joint[n].ToString();
                }

                if (Convert.ToDecimal(manhole[n]) != 0)
                {
                    workSheet.Cells[n + 3, "I"] = manhole[n].ToString();
                    workSheet.Cells[n + 3, "I"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "I"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "I"] = manhole[n].ToString();
                }

                if (Convert.ToDecimal(alligator_crack[n]) != 0)
                {
                    workSheet.Cells[n + 3, "J"] = alligator_crack[n].ToString();
                    workSheet.Cells[n + 3, "J"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "J"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "J"] = alligator_crack[n].ToString();
                }


                if (Convert.ToDecimal(patch[n]) != 0)
                {
                    workSheet.Cells[n + 3, "K"] = patch[n].ToString();
                    workSheet.Cells[n + 3, "K"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "K"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "K"] = patch[n].ToString();
                }

                if (Convert.ToDecimal(white_line[n]) != 0)
                {
                    workSheet.Cells[n + 3, "L"] = white_line[n].ToString();
                    workSheet.Cells[n + 3, "L"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "L"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "L"] = white_line[n].ToString();
                }

                if (Convert.ToDecimal(long_white_line[n]) != 0)
                {
                    workSheet.Cells[n + 3, "M"] = long_white_line[n].ToString();
                    workSheet.Cells[n + 3, "M"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "M"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "M"] = long_white_line[n].ToString();
                }

                if (Convert.ToDecimal(grating[n]) != 0)
                {
                    workSheet.Cells[n + 3, "N"] = grating[n].ToString();
                    workSheet.Cells[n + 3, "N"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "N"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "N"] = grating[n].ToString();
                }

                if (Convert.ToDecimal(linear_crack_p[n]) != 0)
                {
                    workSheet.Cells[n + 3, "O"] = linear_crack_p[n].ToString();
                    workSheet.Cells[n + 3, "O"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "O"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "O"] = linear_crack_p[n].ToString();
                }

                if (Convert.ToDecimal(alligator_crack_p[n]) != 0)
                {
                    workSheet.Cells[n + 3, "P"] = alligator_crack_p[n].ToString();
                    workSheet.Cells[n + 3, "P"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "P"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "P"] = alligator_crack_p[n].ToString();
                }

                if (Convert.ToDecimal(recess[n]) != 0)
                {
                    workSheet.Cells[n + 3, "Q"] = recess[n].ToString();
                    workSheet.Cells[n + 3, "Q"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "Q"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "Q"] = recess[n].ToString();
                }

                if (Convert.ToDecimal(rutting[n]) != 0)
                {
                    workSheet.Cells[n + 3, "R"] = rutting[n].ToString();
                    workSheet.Cells[n + 3, "R"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "R"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "R"] = rutting[n].ToString();
                }

                if (Convert.ToDecimal(pothole[n]) != 0)
                {
                    workSheet.Cells[n + 3, "S"] = pothole[n].ToString();
                    workSheet.Cells[n + 3, "S"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "S"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "S"] = pothole[n].ToString();
                }

                if (Convert.ToDecimal(construction_j[n]) != 0)
                {
                    workSheet.Cells[n + 3, "T"] = construction_j[n].ToString();
                    workSheet.Cells[n + 3, "T"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "T"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "T"] = construction_j[n].ToString();
                }

                if (Convert.ToDecimal(p_longitudinal[n]) != 0)
                {
                    workSheet.Cells[n + 3, "U"] = p_longitudinal[n].ToString();
                    workSheet.Cells[n + 3, "U"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "U"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "U"] = p_longitudinal[n].ToString();
                }

                if (Convert.ToDecimal(p_transverse[n]) != 0)
                {
                    workSheet.Cells[n + 3, "V"] = p_transverse[n].ToString();
                    workSheet.Cells[n + 3, "V"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "V"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "V"] = p_transverse[n].ToString();
                }

                if (Convert.ToDecimal(w_line_crack[n]) != 0)
                {
                    workSheet.Cells[n + 3, "W"] = w_line_crack[n].ToString();
                    workSheet.Cells[n + 3, "W"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "W"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "W"] = w_line_crack[n].ToString();
                }

                if (Convert.ToDecimal(w_alligator_crack[n]) != 0)
                {
                    workSheet.Cells[n + 3, "X"] = w_alligator_crack[n].ToString();
                    workSheet.Cells[n + 3, "X"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "X"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "X"] = w_alligator_crack[n].ToString();
                }

                if (Convert.ToDecimal(y_line_crack[n]) != 0)
                {
                    workSheet.Cells[n + 3, "Y"] = y_line_crack[n].ToString();
                    workSheet.Cells[n + 3, "Y"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "Y"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "Y"] = y_line_crack[n].ToString();
                }


                if (Convert.ToDecimal(y_alligator_crack[n]) != 0)
                {
                    workSheet.Cells[n + 3, "Z"] = y_alligator_crack[n].ToString();
                    workSheet.Cells[n + 3, "Z"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    workSheet.Cells[n + 3, "Z"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                }
                else
                {
                    workSheet.Cells[n + 3, "Z"] = y_alligator_crack[n].ToString();
                }


                Console.WriteLine(n);
            }

            //******** Formula to get the sum of all the non zero values (For DL data)
            workSheet.Cells[totalNumberOfrows + 3, "G"].value2 = "=Sum(G2:G" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "H"].value2 = "=Sum(H2:H" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "I"].value2 = "=Sum(I2:I" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "J"].value2 = "=Sum(J2:J" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "K"].value2 = "=Sum(K2:K" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "L"].value2 = "=Sum(L2:L" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "M"].value2 = "=Sum(M2:M" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "N"].value2 = "=Sum(N2:N" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "O"].value2 = "=Sum(O2:O" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "P"].value2 = "=Sum(P2:P" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "Q"].value2 = "=Sum(Q2:Q" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "R"].value2 = "=Sum(R2:R" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "S"].value2 = "=Sum(S2:S" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "T"].value2 = "=Sum(T2:T" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "U"].value2 = "=Sum(U2:U" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "V"].value2 = "=Sum(V2:V" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "W"].value2 = "=Sum(W2:W" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "X"].value2 = "=Sum(X2:X" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "Y"].value2 = "=Sum(Y2:Y" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";
            workSheet.Cells[totalNumberOfrows + 3, "Z"].value2 = "=Sum(Z2:Z" + Convert.ToDecimal(totalNumberOfrows + 2) + ")";

            //****** Auto fit all the Ranges once data is written
            V2PcsvRange.Columns.AutoFit();
            DatRange.Columns.AutoFit();
            DLRange.Columns.AutoFit();
            DMSDataRange.Columns.AutoFit();

            //************** Range for formula at the bottom and its formatting and color (DL data summary)
            int startColumn4 = 7;
            int startRow4 = totalNumberOfrows + 3;
            Microsoft.Office.Interop.Excel.Range startCell4 = workSheet.Cells[startRow4, startColumn4];

            int endColumn4 = 26;
            int endRow4 = totalNumberOfrows + 3;
            Microsoft.Office.Interop.Excel.Range endCell4 = workSheet.Cells[endRow4, endColumn4];

            Microsoft.Office.Interop.Excel.Range FormulaRange = workSheet.Range[startCell4, endCell4];
            FormulaRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.PowderBlue);

            FormulaRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            FormulaRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            FormulaRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            FormulaRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            FormulaRange.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);

            FormulaRange.Columns.AutoFit();

            //*********** Addintion of sheet two for all the reference values of Deep learning data
            Microsoft.Office.Interop.Excel.Worksheet workSheet2 = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            workSheet2.Name = "Deep_Learning_Column_Names";

            //********** Writing reference values for 0, 1, 2, ..........
            workSheet2.Cells[1, "A"] = " 0: linear crack      線状ひび割れ";
            workSheet2.Cells[2, "A"] = " 1: joint             ジョイント";
            workSheet2.Cells[3, "A"] = " 2: manhole           マンホール";
            workSheet2.Cells[4, "A"] = " 3: alligator crack   亀甲状ひび割れ";
            workSheet2.Cells[5, "A"] = " 4: patch             パッチング";
            workSheet2.Cells[6, "A"] = " 5: white line        短い白線";
            workSheet2.Cells[7, "A"] = " 6: long white_line   長い白線";
            workSheet2.Cells[8, "A"] = " 7: grating           グレーチング";
            workSheet2.Cells[9, "A"] = " 8: linear crack_p    線状ひび割れ補修跡";
            workSheet2.Cells[10, "A"] = " 9: alligator crack_p 亀甲状ひび割れ補修跡";
            workSheet2.Cells[11, "A"] = "10: recess            剥離・くぼみ";
            workSheet2.Cells[12, "A"] = "11: rutting           わだち掘れ";
            workSheet2.Cells[13, "A"] = "12: pothole           ポットホール";
            workSheet2.Cells[14, "A"] = "13: construction_j    施工上の継ぎ目";
            workSheet2.Cells[15, "A"] = "14: p_longitudinal    パッチング縦";
            workSheet2.Cells[16, "A"] = "15: p_transverse      パッチング横";
            workSheet2.Cells[17, "A"] = "16: w_line_crack      白線上の線状ひび割れ";
            workSheet2.Cells[18, "A"] = "17: w_alligator_crack 白線上の亀甲状ひび割れ";
            workSheet2.Cells[19, "A"] = "18: y_line_crack      黄線上の線状ひび割れ";
            workSheet2.Cells[20, "A"] = "19: y_alligator_crack 黄線上の亀甲状ひび割れ";

            //********** Adding reference to all the vales calulated in the first sheet
            workSheet2.Cells[1, "B"].value2 = "=Detailed_Result!G" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[2, "B"].value2 = "=Detailed_Result!H" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[3, "B"].value2 = "=Detailed_Result!I" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[4, "B"].value2 = "=Detailed_Result!J" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[5, "B"].value2 = "=Detailed_Result!K" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[6, "B"].value2 = "=Detailed_Result!L" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[7, "B"].value2 = "=Detailed_Result!M" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[8, "B"].value2 = "=Detailed_Result!N" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[9, "B"].value2 = "=Detailed_Result!O" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[10, "B"].value2 = "=Detailed_Result!P" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[11, "B"].value2 = "=Detailed_Result!Q" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[12, "B"].value2 = "=Detailed_Result!R" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[13, "B"].value2 = "=Detailed_Result!S" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[14, "B"].value2 = "=Detailed_Result!T" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[15, "B"].value2 = "=Detailed_Result!U" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[16, "B"].value2 = "=Detailed_Result!V" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[17, "B"].value2 = "=Detailed_Result!W" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[18, "B"].value2 = "=Detailed_Result!X" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[19, "B"].value2 = "=Detailed_Result!Y" + Convert.ToDecimal(totalNumberOfrows + 3);
            workSheet2.Cells[20, "B"].value2 = "=Detailed_Result!Z" + Convert.ToDecimal(totalNumberOfrows + 3);

            //*********** Range for all the reference values of DL and its formatting
            int startColumn5 = 1;
            int startRow5 = 1;
            Microsoft.Office.Interop.Excel.Range startCell5 = workSheet2.Cells[startRow5, startColumn5];

            int endColumn5 = 1;
            int endRow5 = 20;
            Microsoft.Office.Interop.Excel.Range endCell5 = workSheet2.Cells[endRow5, endColumn5];

            Microsoft.Office.Interop.Excel.Range Deep_Learning_Column_Names = workSheet2.Range[startCell5, endCell5];
            Deep_Learning_Column_Names.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Azure);

            Deep_Learning_Column_Names.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            Deep_Learning_Column_Names.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            Deep_Learning_Column_Names.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            Deep_Learning_Column_Names.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            Deep_Learning_Column_Names.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);

            Deep_Learning_Column_Names.Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            Deep_Learning_Column_Names.Columns.AutoFit();

            //************** Range for all the reference values data from first sheet and its formatting
            int startColumn6 = 2;
            int startRow6 = 1;
            Microsoft.Office.Interop.Excel.Range startCell6 = workSheet2.Cells[startRow6, startColumn6];

            int endColumn6 = 2;
            int endRow6 = 20;
            Microsoft.Office.Interop.Excel.Range endCell6 = workSheet2.Cells[endRow6, endColumn6];

            Microsoft.Office.Interop.Excel.Range Deep_Learning_Column_LinekdData = workSheet2.Range[startCell6, endCell6];
            Deep_Learning_Column_LinekdData.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

            Deep_Learning_Column_LinekdData.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            Deep_Learning_Column_LinekdData.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            Deep_Learning_Column_LinekdData.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            Deep_Learning_Column_LinekdData.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);
            Deep_Learning_Column_LinekdData.Borders.get_Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = (Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous);

            Deep_Learning_Column_LinekdData.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            Deep_Learning_Column_LinekdData.Columns.AutoFit();

            //****** Moving the sheets so that while opening first sheet should be displayed
            wb.Sheets["Deep_Learning_Column_Names"].Move(wb.Sheets[1]);
            wb.Sheets["Detailed_Result"].Move(wb.Sheets[1]);

            //********* Saving the sheet and workbook
            wb.SaveAs(folderPath + "Output_Final_DetailedWithColor" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls");
            //********* Clearing and disposing the excel object
            ClearAllEndAll(excelWrite, workSheet);
            //********** Complete message method
            Defeat();
        }


        /// <summary>
        /// The loadCsvFile
        /// Reads the DAT fil and adds the data into list and removes 
        /// the first two elemets of list as they are not required 
        /// </summary>
        /// <param name="filePath">The filePath<see cref="string"/></param>
        public void loadCsvFile(string filePath)
        {
            var reader = new StreamReader(File.OpenRead(filePath));
            List<string> searchList = new List<string>();
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(',');

                linear_crack.Add(values[1]);
                joint.Add(values[2]);
                manhole.Add(values[3]);
                alligator_crack.Add(values[4]);
                patch.Add(values[5]);
                white_line.Add(values[6]);
                long_white_line.Add(values[7]);
                grating.Add(values[8]);
                linear_crack_p.Add(values[9]);
                alligator_crack_p.Add(values[10]);
                recess.Add(values[11]);
                rutting.Add(values[12]);
                pothole.Add(values[13]);
                construction_j.Add(values[14]);
                p_longitudinal.Add(values[15]);
                p_transverse.Add(values[16]);
                w_line_crack.Add(values[17]);
                w_alligator_crack.Add(values[18]);
                y_line_crack.Add(values[19]);
                y_alligator_crack.Add(values[20]);
            }

            //*** removed first row
            linear_crack.RemoveAt(0);
            joint.RemoveAt(0);
            manhole.RemoveAt(0);
            alligator_crack.RemoveAt(0);
            patch.RemoveAt(0);
            white_line.RemoveAt(0);
            long_white_line.RemoveAt(0);
            grating.RemoveAt(0);
            linear_crack_p.RemoveAt(0);
            alligator_crack_p.RemoveAt(0);
            recess.RemoveAt(0);
            rutting.RemoveAt(0);
            pothole.RemoveAt(0);
            construction_j.RemoveAt(0);
            p_longitudinal.RemoveAt(0);
            p_transverse.RemoveAt(0);
            w_line_crack.RemoveAt(0);
            w_alligator_crack.RemoveAt(0);
            y_line_crack.RemoveAt(0);
            y_alligator_crack.RemoveAt(0);

            //*** removed first row again, so that first two rows are removed
            linear_crack.RemoveAt(0);
            joint.RemoveAt(0);
            manhole.RemoveAt(0);
            alligator_crack.RemoveAt(0);
            patch.RemoveAt(0);
            white_line.RemoveAt(0);
            long_white_line.RemoveAt(0);
            grating.RemoveAt(0);
            linear_crack_p.RemoveAt(0);
            alligator_crack_p.RemoveAt(0);
            recess.RemoveAt(0);
            rutting.RemoveAt(0);
            pothole.RemoveAt(0);
            construction_j.RemoveAt(0);
            p_longitudinal.RemoveAt(0);
            p_transverse.RemoveAt(0);
            w_line_crack.RemoveAt(0);
            w_alligator_crack.RemoveAt(0);
            y_line_crack.RemoveAt(0);
            y_alligator_crack.RemoveAt(0);
        }


        /// <summary>
        /// The Defeat method
        /// Displays the message after completion of the process
        /// </summary>
        private void Defeat()
        {
            MessageBox.Show("プロセスは正常に完了しました。 ");
            this.Close();
        }

        /// <summary>
        /// The ClearAllEndAll method
        /// Disposes the excel object and handles the garbage collection implecitly
        /// </summary>
        /// <param name="excelObj">The excelObj<see cref="Microsoft.Office.Interop.Excel.Application"/></param>
        /// <param name="sheetObj">The sheetObj<see cref="Microsoft.Office.Interop.Excel._Worksheet"/></param>
        private void ClearAllEndAll(Microsoft.Office.Interop.Excel.Application excelObj, Microsoft.Office.Interop.Excel._Worksheet sheetObj)
        {
            // Quit Excel application
            excelObj.Quit();

            // Release COM objects (very important!)
            if (excelObj != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelObj);

            if (sheetObj != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheetObj);

            // Empty variables
            excelObj = null;
            sheetObj = null;

            // Force garbage collector cleaning
            GC.Collect();
        }

        /// <summary>
        /// Defines the <see cref="CsvData" />
        /// Defines the data present in V2P file
        /// </summary>
        internal class CsvData
        {
            /// <summary>
            /// Gets or sets the ImageFileName
            /// </summary>
            [CsvColumn(Name = "ImageFileName", FieldIndex = 1)]
            public string ImageFileName { get; set; }

            /// <summary>
            /// Gets or sets the ShootDate
            /// </summary>
            [CsvColumn(Name = "ShootDate", FieldIndex = 2, OutputFormat = "yyyy/MM/dd")]
            public DateTime ShootDate { get; set; }

            /// <summary>
            /// Gets or sets the TimeStamp
            /// </summary>
            [CsvColumn(Name = "TimeStamp", FieldIndex = 3, OutputFormat = "HH:mm:ss")]
            public DateTime TimeStamp { get; set; }

            /// <summary>
            /// Gets or sets the Latitude
            /// </summary>
            [CsvColumn(Name = "Latitude", FieldIndex = 4, CanBeNull = false, OutputFormat = "M")]
            public Decimal Latitude { get; set; }

            /// <summary>
            /// Gets or sets the Longitude
            /// </summary>
            [CsvColumn(Name = "Longitude", FieldIndex = 5, CanBeNull = false, OutputFormat = "M")]
            public Decimal Longitude { get; set; }

            /// <summary>
            /// Gets or sets the horizontalAccuracy
            /// </summary>
            [CsvColumn(Name = "speed", FieldIndex = 6)]
            public Decimal horizontalAccuracy { get; set; }

            /// <summary>
            /// Gets or sets the verticalAccuracy
            /// </summary>
            [CsvColumn(Name = "verticalAccuracy", FieldIndex = 7)]
            public Decimal verticalAccuracy { get; set; }

            /// <summary>
            /// Gets or sets the course
            /// </summary>
            [CsvColumn(Name = "course", FieldIndex = 8)]
            public Decimal course { get; set; }

            /// <summary>
            /// Gets or sets the speed
            /// </summary>
            [CsvColumn(Name = "horizontalAccuracy", FieldIndex = 9)]
            public Decimal speed { get; set; }
        }

        /// <summary>
        /// Defines the CsvFileDescription (V2P file discription)
        /// </summary>
        internal CsvFileDescription CsvFileDescription = new CsvFileDescription
        {
            SeparatorChar = ',',
            FirstLineHasColumnNames = false,
            EnforceCsvColumnAttribute = true,
            TextEncoding = Encoding.UTF8,
            UseFieldIndexForReadingData = true
        };

        /// <summary>
        /// Defines the <see cref="DatFileData" />
        /// Defines the data present in DAT file 
        /// </summary>
        internal class DatFileData
        {
            /// <summary>
            /// Gets or sets the localTimeStamp
            /// </summary>
            [CsvColumn(Name = "localTimeStamp", FieldIndex = 1)]
            public DateTime localTimeStamp { get; set; }

            /// <summary>
            /// Gets or sets the gpsTimeStamp
            /// </summary>
            [CsvColumn(Name = "gpsTimeStamp", FieldIndex = 2)]
            public DateTime gpsTimeStamp { get; set; }

            /// <summary>
            /// Gets or sets the latitude
            /// </summary>
            [CsvColumn(Name = "latitude", FieldIndex = 3)]
            public Decimal latitude { get; set; }

            /// <summary>
            /// Gets or sets the longitude
            /// </summary>
            [CsvColumn(Name = "longitude", FieldIndex = 4)]
            public Decimal longitude { get; set; }

            /// <summary>
            /// Gets or sets the altitude
            /// </summary>
            [CsvColumn(Name = "altitude", FieldIndex = 5)]
            public Decimal altitude { get; set; }

            /// <summary>
            /// Gets or sets the horizontalAccuracy
            /// </summary>
            [CsvColumn(Name = "horizontalAccuracy", FieldIndex = 6)]
            public Decimal horizontalAccuracy { get; set; }

            /// <summary>
            /// Gets or sets the verticalAccuracy
            /// </summary>
            [CsvColumn(Name = "verticalAccuracy", FieldIndex = 7)]
            public Decimal verticalAccuracy { get; set; }

            /// <summary>
            /// Gets or sets the speed
            /// </summary>
            [CsvColumn(Name = "speed", FieldIndex = 8)]
            public Decimal speed { get; set; }

            /// <summary>
            /// Gets or sets the course
            /// </summary>
            [CsvColumn(Name = "course", FieldIndex = 9)]
            public Decimal course { get; set; }
        }

        /// <summary>
        /// Defines the datInputFileDescription (DAT file discription)
        /// </summary>
        internal CsvFileDescription datInputFileDescription = new CsvFileDescription
        {
            SeparatorChar = ',',
            FirstLineHasColumnNames = false,
            EnforceCsvColumnAttribute = true,
            TextEncoding = Encoding.UTF8,
            UseFieldIndexForReadingData = true,
            IgnoreUnknownColumns = true

        };

        /// <summary>
        /// Defines the <see cref="DLCsvFileData" />
        /// Defines the data present in Deep learning output CSV file
        /// </summary>
        internal class DLCsvFileData
        {
            /// <summary>
            /// Gets or sets the ImageFileName
            /// </summary>
            [CsvColumn(Name = "ImageFileName", FieldIndex = 1)]
            public string ImageFileName { get; set; }

            /// <summary>
            /// Gets or sets the linearCrack
            /// </summary>
            [CsvColumn(Name = "linearCrack", FieldIndex = 2)]
            public Decimal linearCrack { get; set; }

            /// <summary>
            /// Gets or sets the joint
            /// </summary>
            [CsvColumn(Name = "joint", FieldIndex = 3)]
            public Decimal joint { get; set; }

            /// <summary>
            /// Gets or sets the manhole
            /// </summary>
            [CsvColumn(Name = "manhole", FieldIndex = 4)]
            public Decimal manhole { get; set; }

            /// <summary>
            /// Gets or sets the alligatorCrack
            /// </summary>
            [CsvColumn(Name = "alligatorCrack", FieldIndex = 5)]
            public Decimal alligatorCrack { get; set; }

            /// <summary>
            /// Gets or sets the patch
            /// </summary>
            [CsvColumn(Name = "patch", FieldIndex = 6)]
            public Decimal patch { get; set; }

            /// <summary>
            /// Gets or sets the whiteLine
            /// </summary>
            [CsvColumn(Name = "whiteLine", FieldIndex = 7)]
            public Decimal whiteLine { get; set; }

            /// <summary>
            /// Gets or sets the longWhiteLine
            /// </summary>
            [CsvColumn(Name = "longWhiteLine", FieldIndex = 8)]
            public Decimal longWhiteLine { get; set; }

            /// <summary>
            /// Gets or sets the grating
            /// </summary>
            [CsvColumn(Name = "grating", FieldIndex = 9)]
            public Decimal grating { get; set; }

            /// <summary>
            /// Gets or sets the linearCrack_p
            /// </summary>
            [CsvColumn(Name = "linearCrack_p", FieldIndex = 10)]
            public Decimal linearCrack_p { get; set; }

            /// <summary>
            /// Gets or sets the alligatorCrack_p
            /// </summary>
            [CsvColumn(Name = "alligatorCrack_p", FieldIndex = 11)]
            public Decimal alligatorCrack_p { get; set; }

            /// <summary>
            /// Gets or sets the recess
            /// </summary>
            [CsvColumn(Name = "recess", FieldIndex = 12)]
            public Decimal recess { get; set; }

            /// <summary>
            /// Gets or sets the rutting
            /// </summary>
            [CsvColumn(Name = "rutting", FieldIndex = 13)]
            public Decimal rutting { get; set; }

            /// <summary>
            /// Gets or sets the pothole
            /// </summary>
            [CsvColumn(Name = "pothole", FieldIndex = 14)]
            public Decimal pothole { get; set; }

            /// <summary>
            /// Gets or sets the construction_j
            /// </summary>
            [CsvColumn(Name = "construction_j", FieldIndex = 15)]
            public Decimal construction_j { get; set; }

            /// <summary>
            /// Gets or sets the p_longitudinal
            /// </summary>
            [CsvColumn(Name = "p_longitudinal", FieldIndex = 16)]
            public Decimal p_longitudinal { get; set; }

            /// <summary>
            /// Gets or sets the p_transverse
            /// </summary>
            [CsvColumn(Name = "p_transverse", FieldIndex = 17)]
            public Decimal p_transverse { get; set; }

            /// <summary>
            /// Gets or sets the w_lineCrack
            /// </summary>
            [CsvColumn(Name = "w_lineCrack", FieldIndex = 18)]
            public Decimal w_lineCrack { get; set; }

            /// <summary>
            /// Gets or sets the w_alligatorCrack
            /// </summary>
            [CsvColumn(Name = "w_alligatorCrack", FieldIndex = 19)]
            public Decimal w_alligatorCrack { get; set; }

            /// <summary>
            /// Gets or sets the y_lineCrack
            /// </summary>
            [CsvColumn(Name = "y_lineCrack", FieldIndex = 20)]
            public Decimal y_lineCrack { get; set; }

            /// <summary>
            /// Gets or sets the y_alligatorCrack
            /// </summary>
            [CsvColumn(Name = "y_alligatorCrack", FieldIndex = 21)]
            public Decimal y_alligatorCrack { get; set; }
        }

        /// <summary>
        /// Defines the DLInputFileDescription (Deep Learning output csv file discription)
        /// </summary>
        internal CsvFileDescription DLInputFileDescription = new CsvFileDescription
        {
            SeparatorChar = ',',
            FirstLineHasColumnNames = false,
            EnforceCsvColumnAttribute = true,
            TextEncoding = Encoding.UTF8,
            UseFieldIndexForReadingData = true,
        };
    }
}
