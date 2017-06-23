using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace M3_Preprocess_Runner
{
    public partial class Form1 : Form
    {
        BackgroundWorker Worker;
        M3TargetStruct M3TargetData;
        public delegate void MessageHandlerDelegate(string Key, string MessageToHandle);                // Step 1 for event handling, declare a delegate
        public event MessageHandlerDelegate MessageToHandleEvent;                                       // Step 2 for event handling, declare an event - see M3Class
        private string BatchHoursText= "";
        private string BatchAmountText = "";

        public Form1()
        {
            this.MessageToHandleEvent += new MessageHandlerDelegate(Form1_MessageHandler);              // Step 3 for event handling, get an instance of this class and assign to the event the delegate of the handler method
            M3TargetData = new M3TargetStruct();
            M3TargetData.DSN = "Millennium";
            M3TargetData.Username = "TSR";
            M3TargetData.Password = "Tricore2#";
            M3TargetData.Company = "CARD2"; //"IMPA"; //"CHOI"; //"CARD2"; //"XRP"; // "KAFL"; // "CROW"; //XRP has no calendar

            InitializeComponent();
            this.Text = "M3 Precalc C# Script";
            this.textBox1.ReadOnly = true;
            this.textBox1.Text =  "DSN          : " + "\r\n" + "User         : " + "\r\n" + "\r\n";
            this.textBox1.Text += "Company      : " + "\r\n" + "Check Date   : " + "\r\n" + "\r\n";
            this.textBox1.Text += "Total Hours  : " + "\r\n" + "Total Amount : " + "\r\n" + "\r\n";
            this.textBox1.Text += "Check Count  : " + "\r\n" + "\r\n";
            this.textBox1.Text += "Reports Count: " + "\r\n" + "\r\n" + "Status       : " + "\r\n" + "\r\n";
            this.textBox1.SelectionStart = 0;
            this.textBox1.SelectionLength = 0;

            if (MessageBox.Show(this, "This program will...\r\n\r\n\t\u2022 Sync with PayEntry\r\n\t\u2022 Pre-calc payroll\r\n\t\u2022 Run preprocess reports\r\n\r\nfor: " + M3TargetData.Company + "\r\n\r\n\r\nContinue?", "M3 Preprocess Report Runner", MessageBoxButtons.OKCancel) != System.Windows.Forms.DialogResult.OK)
            {
                this.Close();
                return;
            }

            Worker = new BackgroundWorker();

            //Worker.DoWork += new DoWorkEventHandler( (object s, DoWorkEventArgs e) => { });
            //Worker.DoWork += ((object s, DoWorkEventArgs e) => { });

            Worker.DoWork += new DoWorkEventHandler(WorkerSetM3Class);
            Worker.ProgressChanged += new ProgressChangedEventHandler(WorkerProgressChanged);
            Worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(WorkerRunWorkerCompleted);
            Worker.WorkerReportsProgress = true;
            Worker.WorkerSupportsCancellation = true;
            Worker.RunWorkerAsync(M3TargetData); //  can pass any data type
        }

        // *** EVENT HANDLER - THREAD-SAFE ***
        public void Form1_MessageHandler(string Key, string MessageToHandle)                        // Step 5 for event handling, write a thread-safe handler
        {
            if (this.InvokeRequired) // <-- important technique to service calls from other threads
            {
                MessageHandlerDelegate textWriter = new MessageHandlerDelegate(Form1_HandleMessages);
                this.Invoke(textWriter, new object[] { Key, MessageToHandle });
            }
            else
            {
                this.Form1_HandleMessages(Key, MessageToHandle);
            }
        }

        private void Form1_HandleMessages(string Key, string MessageToHandle)
        {
            string Message = MessageToHandle.Replace("\r\n", "\r\n" + "               ").ToString();
            string TextToUpdate;
            string TextToFind1, TextToFind2;
            int longerLen = 0;
            string numPadding = "";

            TextToUpdate = this.textBox1.Text;
            switch (Key)
            {
                case "Status":     TextToFind1 = "Status       : "; TextToFind2 = "\r\nDSN          : "    ; break;
                case "DSN":        TextToFind1 = "DSN          : "; TextToFind2 = "\r\nUser         : "    ; break;
                case "User":       TextToFind1 = "User         : "; TextToFind2 = "\r\n\r\nCompany      : "; break;
                case "Company":    TextToFind1 = "Company      : "; TextToFind2 = "\r\nCheck Date   : "    ; break;
                //case "Calendar": TextToFind1 = "Calendar     : "; TextToFind2 = "Status       : "        ; break;
                case "Check Date": TextToFind1 = "Check Date   : "; TextToFind2 = "\r\n\r\nTotal Hours  : "; break;
                case "Hours":      TextToFind1 = "Total Hours  : "; TextToFind2 = "\r\nTotal Amount : ";
                    BatchHoursText = Message;
                    longerLen = (BatchAmountText.Length > BatchHoursText.Length) ? BatchAmountText.Length : BatchHoursText.Length;
                    numPadding = new String(' ', longerLen);
                    Message = (numPadding + Message.ToString()).Substring(Message.ToString().Length);
                    break;
                case "Amount":     TextToFind1 = "Total Amount : "; TextToFind2 = "\r\n\r\nCheck Count  : ";
                    BatchAmountText = Message;
                    longerLen = (BatchAmountText.Length > BatchHoursText.Length) ? BatchAmountText.Length : BatchHoursText.Length;
                    numPadding = new String(' ', longerLen);
                    Message = (numPadding + Message.ToString()).Substring(Message.ToString().Length);
                    break;
                case "Check Count": TextToFind1 = "Check Count  : "; TextToFind2 = "\r\n\r\nReports Count: "; break;
                case "Reports Count": TextToFind1 = "Reports Count: "; TextToFind2 = "\r\n\r\nStatus       : "; break;
                default:           TextToFind1 = "Status       : "; TextToFind2 = ""                       ; break;
            }

            StringBuilder sb = new StringBuilder(TextToUpdate);

            if (TextToFind1 == "Status       : " && TextToUpdate.Contains("               "))
            {
                TextToFind1 = "               ";
            }

            if (TextToFind1 == "               " || TextToFind1 == "Status       : ")
            {
                sb.Insert(TextToUpdate.LastIndexOf(TextToFind1) + TextToFind1.Length, Message.ToString() + "\r\n" + "               ");
            }
            else
            {
                sb.Remove(TextToUpdate.IndexOf(TextToFind1) + TextToFind1.Length, TextToUpdate.IndexOf(TextToFind2) - (TextToUpdate.IndexOf(TextToFind1) + TextToFind1.Length));
                sb.Insert(TextToUpdate.IndexOf(TextToFind1) + TextToFind1.Length, Message.ToString());
            }
            this.textBox1.Text = sb.ToString();
        }

        void WorkerSetM3Class(object sender, DoWorkEventArgs e)
        {
            PEClass PE;
            PE = new PEClass();
            PE.MessageToPublishEvent += new PEClass.PublishMessageDelegate(Form1_MessageHandler);       // Step 3 for event handling, assign a handler
            PE.LoginAndSync();

            if (MessageBox.Show("PayEntry should be synchronizing.\r\n\r\nPlease check the M3 job queue and click OK when the sync is complete.", "M3 Preprocess Report Runner", MessageBoxButtons.OKCancel) != System.Windows.Forms.DialogResult.OK)
            {
                Application.Exit();
                return;
            }

            M3Class M3;
            string d, u, p, c, m3msg, t, f, n;
            var M3Target = (M3TargetStruct)M3TargetData;
            d = M3Target.DSN; u = M3Target.Username; p = M3Target.Password; c = M3Target.Company;

            this.MessageToHandleEvent("", "Loading the M3 system...");
            try
            {
                M3 = new M3Class(); //not async
                M3.M3Messages.TryGetValue("SetM3System", out m3msg); // this message appears after the async messages
                this.MessageToHandleEvent("", m3msg);
            }
            catch (Exception err)
            {
                this.MessageToHandleEvent("", "Error loading the M3 system");
                this.MessageToHandleEvent("", err.Message);
                return;
            }

            M3.MessageToPublishEvent += new M3Class.PublishMessageDelegate(Form1_MessageHandler);       // Step 3 for event handling, assign a handler

            //M3.SetDSNAsync(d);

            //M3.SetUserAsync(u, p);
            //M3.SetUserAsync(d, u, p);

            //M3.SetCompanyAsync(c);
            //M3.SetCompanyAsync(d, u, p, c).Wait();

            //M3.LoadCompanyCalendarAsync();
            //M3.LoadCompanyCalendarAsync(c);
            //M3.LoadCompanyCalendarAsync(d, u, p, c).Wait();

            //M3.AggregateBatchTotalsAsync();
            //M3.AggregateBatchTotalsAsync(c); //.Wait();
            //M3.AggregateBatchTotalsAsync(d, u, p, c).Wait();

            //M3.PreCalcAsync();
            //M3.PreCalcAsync(c);
            //M3.PreCalcAsync(d, u, p, c).Wait();

            t = @"Reporting/Preprocess Reports"; //@"Reporting/PreprocessRpts/PreProcByLoc"; // @"Reporting/Preprocess Reports";
            f = ""; // @"\\HOST3\Millennium Shared Files\Temp"; // ""; //@"\\HOST3\Millennium Shared Files\Temp"; // @"\\Ed\Desktop\CARD2 PreProcess";
            n = "";
            //M3.RunReportsFromTreeAsync(t, f, n);  // "M3 tree folder, UNC targer folder, pattern for filter
            //M3.RunReportsFromTreeAsync("CARD2", t, f, n);
            //M3.RunReportsFromTreeAsync(d, u, p, c, t, f, n);

            //M3.RunReportsFromTreeWithPreCalcAsync(t, f, n);  // "M3 tree folder, UNC targer folder, pattern for filter
            //M3.RunReportsFromTreeWithPreCalcAsync("CARD2", t, f, n);
            //M3.RunReportsFromTreeWithPreCalcAsync(d, u, p, c, t, f, n);
            M3.RunReportsFromTreeWithPreCalcAsync(d, u, p, c, t, f, n).Wait();
            Form1_MessageHandler("Status", "Program complete");
        }

        void WorkerProgressChanged(object sender, ProgressChangedEventArgs e)
        {   // This function fires on the UI thread so it's safe to edit
            // the UI control directly, no funny business with Control.Invoke :)
            //int i = e.ProgressPercentage;
            //this.MessageToHandleEvent("", i + "%");
            //Form1_HandleMessages("", i + "%");
        }

        void WorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
        }
  
        private struct M3TargetStruct
        {
            public string DSN;
            public string Username;
            public string Password;
            public string Company;
            //public string EmployeeId;
            //public string Batch;
        }

    }
}