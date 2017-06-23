using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using M3_Preprocess_Runner;
using System.Threading;
using System.Collections.Concurrent;
using System.Globalization;

/// <summary>
/// This class wraps the M3 object model with asynchronous C# methods that return a task, while publishing status updates through an event.</summary>
/// <remarks>
/// Everything except the initial constructor is asynchronous.
/// M3Class(), which is not async, loads the M3 object model.
/// 
/// Written March 2014 by Ed Pantaleone primarily to run 40 preprocess reports for First Choice Loan Services (FCLS)
/// 
/// SetDSNAsync(d) connects to the specific database
/// 
/// SetUserAsync(u, p) attempts to login a user, but SetDSNAsync(d) must be used first
/// SetUserAsync(d, u, p) connects to the specific database and attempts to login the user - no need to preset the DSN
/// 
/// SetCompanyAsync(c) attempts to load the specific company, but a database connection and user session must exist
/// SetCompanyAsync(d, u, p, c) connects to the specific database, attempts to login the user and to load the company - no prior dependency except class instantiation
/// 
/// LoadCompanyCalendarAsync() loads the company calendar
/// LoadCompanyCalendarAsync(c) loads the calendar for the specified company, resetting the company if necessasry
/// LoadCompanyCalendarAsync(d, u, p, c) connects to the specific database, attempts to login the user, loads the company and attempts to load the calendar - no prior dependency except class instantiation
/// 
/// AggregateBatchTotalsAsync() aggregates open batch totals (hours and amounts) for each batch and each earnings code found
/// AggregateBatchTotalsAsync(c) aggregates open batch totals (hours and amounts) for the company specified, resetting the company if necessary
/// AggregateBatchTotalsAsync(d, u, p, c) connects to the specific database, attempts to login the user, loads the company, loads the calendar and attempts to aggregate open batch totals (hours and amounts) - no prior dependency except class instantiation
/// 
/// PreCalcAsync() loads the company calendar, aggregates open batch totals and pre-calculates all paychecks
/// PreCalcAsync(c) pre-calculates all paychecks for the company specified, resetting the company if necessary (See the 1st overload for the PreCalcAsync description)
/// PreCalcAsync(d, u, p, c) connects to the specific database, attempts to login the user, loads the company, loads the calendar, and attempts to aggregate batch totals and to pre-calcs payroll - no prior dependency except class instantiation
/// 
/// RunReportsFromTreeAsync(t, o, n) runs reports existing in an M3 tree folder using a path override option and a partial-text label or name filter (treefolder, outputUNCPath, pattern);
/// RunReportsFromTreeAsync(c, t, o, n) runs reports existing in an M3 tree folder for the company specified, resetting the company if necessary (See the 1st overload for the RunReportsFromTreeAsync description)
/// RunReportsFromTreeAsync(d, u, p, c, t, o, n) connects to the specific database, attempts to login the user, load the company, and attempts to run reports existing in an M3 tree folder - no prior dependency except class instantiation
/// 
/// RunReportsFromTreeWithPreCalcAsync(t, o, n) pre-calculates all paychecks and runs reports existing in an M3 tree folder using a path override option and a partial-text label or name filter (treefolder, outputUNCPath, pattern);
/// RunReportsFromTreeWithPreCalcAsync(c, t, o, n) pre-calculates all paychecks and runs reports existing in an M3 tree folder for the company specified, resetting the company if necessary (See the 1st overload for the RunReportsFromTreeWithPreCalcAsync description)
/// RunReportsFromTreeWithPreCalcAsync(d, u, p, c, t, o, n) connects to the specific database, attempts to login the user, loads the company, loads the calendar, aggregates batch totals, pre-calcs payroll and attempts to run reports existing in an M3 tree folder - no prior dependency except class instantiation
/// 
/// If a task fails, it cancels any outstanding task and returns
/// </remarks>

namespace M3_Preprocess_Runner
{
    public class M3Class
    {
        public delegate void PublishMessageDelegate(string Key, string MessageToPublish);       // Step 1 for event handling, declare a delegate
        public event PublishMessageDelegate MessageToPublishEvent;                              // Step 2 for event handling, declare an event
        // Step 3 for event handling, assign a handler to listen for this event (usually in another class)
        public MILLSYSTEMLib.System M3System;
        public MILLSYSTEMLib.SUser M3User;
        public MILLCOMPANYLib.Company M3Company;
        public MILLCOMPANYLib.CCalendar M3CompanyCalendar;
        public MILLTREELib.TreeHelper M3TreeHelper;

        public ConcurrentDictionary<string, string> M3Messages;
        public CancellationTokenSource M3ClassCancellationNotice;

        Task SetDSNAsyncTask;
        Task SetUserAsyncTask;
        Task SetCompanyAsyncTask;
        Task PreCalcAsyncTask;
        Task LoadCompanyCalendarAsyncTask;
        Task AggregateBatchTotalsAsyncTask;
        Task RunReportsFromTreeAsyncTask;

        private double allBatchTotalHours = 0;
        private double allBatchTotalAmounts = 0;
        private int preCalcCheckCount = 0;
        public double AllBatchTotalHours { get { return allBatchTotalHours; } }
        public double AllBatchTotalAmounts { get { return allBatchTotalAmounts; } }
        public double PreCalcCheckCount { get { return preCalcCheckCount; } }

        private MILLTREELib.TNode M3FolderNode;
        private MILLTREELib.IMPICollection M3FolderNodeItemsCollection;
        private MILLREPORTLib.RReport M3Report;
        private CultureInfo ci = new CultureInfo("en-us");

        private struct ReportJobParametersStruct
        {
            public string M3CompanyCode;
            public string M3CompanyTreePath;            // the M3 tree path within the company tree, start with "Reporting/..." for ad-hoc or "Company Maintenance/Reporting/... for scheduled reports ("Every Payroll")
            public string TargetOutputUNCPath;          // a network share UNC - "\\HOST3\Millennium Shared Files\Temp"
            public string M3ReportLabelOrTitlePattern;  // leave blank to run all reports
        }

        private static ReportJobParametersStruct ReturnReportJobParametersStruct(string M3CompanyCode, string M3CompanyTreePath, string TargetOutputUNCPath, string M3ReportLabelOrTitlePattern)
        {
            ReportJobParametersStruct NewReportJobParametersStruct = new ReportJobParametersStruct();
            string NewM3CompanyTreePath = M3CompanyTreePath;

            // trim leading or trailing slashes
            while (NewM3CompanyTreePath.Length > 0 && NewM3CompanyTreePath.Substring(0, 1) == "/")
            {
                NewM3CompanyTreePath = NewM3CompanyTreePath.Substring(1);
            }
            while (NewM3CompanyTreePath.Length > 0 && NewM3CompanyTreePath.Substring(NewM3CompanyTreePath.Length - 1, 1) == "/")
            {
                NewM3CompanyTreePath = NewM3CompanyTreePath.Substring(0, NewM3CompanyTreePath.Length - 1);
            }
            NewReportJobParametersStruct.M3CompanyCode = M3CompanyCode;
            NewReportJobParametersStruct.M3CompanyTreePath = "/UIRoot/CompanySets/Companies/" + NewM3CompanyTreePath;
            NewReportJobParametersStruct.TargetOutputUNCPath = TargetOutputUNCPath;
            NewReportJobParametersStruct.M3ReportLabelOrTitlePattern = M3ReportLabelOrTitlePattern;
            return NewReportJobParametersStruct;
        }

        // *** CLASS CONSTRUCTOR ***

        /// <summary>
        /// Wraps the Millennium M3 object model with asynchronous C# where possible.</summary>
        public M3Class()
        {
            M3Messages = new ConcurrentDictionary<string, string>();
            M3ClassCancellationNotice = new CancellationTokenSource();
            SetM3System(); // tags used to identify messages in the dictionary class field, populated by each sub
        }

        private void LogMessage(string keyTag, string MessageToLog)
        {
            if (M3Messages.ContainsKey(keyTag))
            {
                string MessageValue = "";
                M3Messages.TryRemove(keyTag, out MessageValue);
            }
            try
            {
                M3Messages.TryAdd(keyTag, MessageToLog);
            }
            catch (Exception e)
            {
                this.MessageToPublishEvent("Status", "M3 class message dictionary error: " + e.Message);
            }
        }

        /// <summary>
        /// Loads the M3 object model - not async</summary>
        private void SetM3System() // part of the constructor 
        {
            string classMsgKeyTag = "SetM3System";
            string MPIMsgKeyTag = "MPISystem";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
            try
            {
                if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
                M3System = new MILLSYSTEMLib.System();
                LogMessage(classMsgKeyTag, "M3 system loaded");
            }
            catch (Exception e)
            {
                LogMessage(classMsgKeyTag, "Error loading M3");
                LogMessage(MPIMsgKeyTag, e.Message);
                throw new System.ArgumentException(e.Source.ToString(), e.Message); // catch this when calling
            }
        }

        // *** SET DATA SOURCE NAME ***

        /// <summary>
        /// Connects to an M3 database asynchronously</summary>
        /// <param name="dsn"> The data source name is the target M3 database name.</param>
        public async Task SetDSNAsync(string dsn)
        {
            string classMsgKeyTag = "SetDSNAsync";
            string MPIMsgKeyTag = "MPILoadDSN";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }

            SetDSNAsyncTask = Task.Factory.StartNew(() => // if the lambda called a task then use async () =>  async would otherwise still runs synchronously
            {
                try
                {
                    if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
                    M3System.Load(dsn);
                    LogMessage(classMsgKeyTag, "Loaded DSN " + dsn);
                    this.MessageToPublishEvent("DSN", this.M3System.DSN);               // Step 4 for event handling, raise the event
                    this.MessageToPublishEvent("Status", "DSN " + this.M3System.DSN + " loaded");
                }
                catch (Exception e)
                {
                    LogMessage(classMsgKeyTag, "Error loading DSN");
                    LogMessage(MPIMsgKeyTag, e.Message);
                    this.MessageToPublishEvent("Status", "Error loading DSN");
                    this.MessageToPublishEvent("Status", e.Message);
                    this.M3ClassCancellationNotice.Cancel();
                }
            });
            await SetDSNAsyncTask;

            // This method has no return statement, so its return type is Task. http://msdn.microsoft.com/en-us/library/hh524395.aspx#BKMK_TaskReturnType
        }

        // *** USER LOGIN ***

        /// <summary>
        /// Attempts to login a user, but SetDSNAsync(d) must be used first</summary>
        /// <param name="M3UserName"> An M3 username</param>
        /// <param name="M3Password"> The M3 password</param>
        public async Task SetUserAsync(string M3UserName, string M3Password)
        {
            string classMsgKeyTag = "SetUserAsync";
            string MPIMsgKeyTag = "MPILogin";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
            if (!SetDSNAsyncTask.IsCompleted) { SetDSNAsyncTask.Wait(); }

            SetUserAsyncTask = Task.Factory.StartNew(() => // if the lambda called a task then use async () =>  async would otherwise still runs synchronously
            {
                try
                {
                    if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
                    this.MessageToPublishEvent("Status", M3UserName + " signing in...");
                    M3User = (MILLSYSTEMLib.SUser)M3System.Login(M3UserName, M3Password); //returns an M3 Entity, not an M3 SUser, hence the cast
                    LogMessage(classMsgKeyTag, "User " + M3User.userName + " sign-in successful");
                    this.MessageToPublishEvent("User", M3User.userName);
                    this.MessageToPublishEvent("Status", "Sign-in complete");
                }
                catch (Exception e)
                {
                    LogMessage(classMsgKeyTag, M3UserName + " user sign-in error");
                    LogMessage(MPIMsgKeyTag, e.Message);
                    this.MessageToPublishEvent("Status", M3UserName + " user sign-in error");
                    this.MessageToPublishEvent("Status", e.Message);
                    this.M3ClassCancellationNotice.Cancel();
                }
            });
            await SetUserAsyncTask;
        }

        /// <summary>
        /// Connects to the specific database and attempts to login the user - no need to preset the DSN</summary>
        /// <param name="dsn"> The data source name is the target M3 database name.</param>
        /// <param name="M3UserName"> An M3 username</param>
        /// <param name="M3Password"> The M3 password</param>
        public async Task SetUserAsync(string dsn, string M3UserName, string M3Password)
        {
            if (M3System.DSN != dsn)
            {
                SetDSNAsync(dsn).Wait();
            }
            await SetUserAsync(M3UserName, M3Password); // note different function signature
        }

        // *** SET COMPANY ***

        /// <summary>
        /// Attempts to load the specified company - the DSN must be preset</summary>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        public async Task SetCompanyAsync(string M3CompanyCode)
        {
            string classMsgKeyTag = "SetCompanyAsync";
            string MPIMsgKeyTag = "MPICompany";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
            if (M3System.DSN == "" || M3User == null)
            {
                LogMessage(classMsgKeyTag, "Could not load company " + M3CompanyCode);
                LogMessage(MPIMsgKeyTag, "n/a");
                this.MessageToPublishEvent("Status", "Error loading company " + M3CompanyCode);
                this.M3ClassCancellationNotice.Cancel();
                return;
            }
            if (!SetDSNAsyncTask.IsCompleted) { SetDSNAsyncTask.Wait(); }
            if (!SetUserAsyncTask.IsCompleted) { SetUserAsyncTask.Wait(); }

            SetCompanyAsyncTask = Task.Factory.StartNew(() => // if the lambda called a task then use async () =>  async would otherwise still runs synchronously
            {
                try
                {
                    if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
                    this.MessageToPublishEvent("Status", "Loading company " + M3CompanyCode + "...");
                    M3Company = (MILLCOMPANYLib.Company)M3System.GetEntity("MillCompany.Company", M3CompanyCode);
                    if (M3Company == null || M3Company.co != M3CompanyCode) { throw new System.ArgumentException("", ""); }
                    LogMessage(classMsgKeyTag, "Company " + M3Company.co + " loaded");
                    this.MessageToPublishEvent("Company", M3Company.co);
                    this.MessageToPublishEvent("Status", "Company " + M3CompanyCode + " loaded");
                }
                catch (Exception e)
                {
                    LogMessage(classMsgKeyTag, "Could not load company " + M3CompanyCode);
                    LogMessage(MPIMsgKeyTag, e.Message);
                    this.MessageToPublishEvent("Status", "Error loading company " + M3CompanyCode);
                    this.MessageToPublishEvent("Status", e.Message);
                    this.M3ClassCancellationNotice.Cancel();
                }
            });
            await SetCompanyAsyncTask;
        }

        /// <summary>
        /// Connects to the specific database, attempts to login a user and to load the company - no need to preset the DSN</summary>
        /// <param name="dsn"> The data source name is the target M3 database name.</param>
        /// <param name="M3UserName"> An M3 username</param>
        /// <param name="M3Password"> The M3 password</param>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        public async Task SetCompanyAsync(string dsn, string M3UserName, string M3Password, string M3CompanyCode)
        {
            if (M3System.DSN != dsn)
            {
                SetDSNAsync(dsn).Wait();
            }

            if (M3User == null || M3User.userName != M3UserName)
            {
                SetUserAsync(M3UserName, M3Password).Wait();
            }
            await SetCompanyAsync(M3CompanyCode);
        }

        // *** LOAD CALENDAR ***

        /// <summary>
        /// Loads the calendar of the current company, but requires an existing M3 user session</summary>
        public async Task LoadCompanyCalendarAsync()
        {
            string classMsgKeyTag = "LoadCompanyCalendarAsync";
            string MPIMsgKeyTag = "MPILoadCompanyCalendar";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }

            LoadCompanyCalendarAsyncTask = Task.Factory.StartNew(() => // if the lambda called a task then use async () =>  async would otherwise still runs synchronously
            {
                try
                {
                    if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
                    this.MessageToPublishEvent("Status", "Loading company calendar...");
                    //M3CompanyCalendar = M3Company.GetEntity("MillCompany.CCalendar", M3Company.calendarId); // two ways to get the calendar
                    M3CompanyCalendar = (MILLCOMPANYLib.CCalendar)M3Company.Calendar[M3Company.calendarId];
                    if (M3CompanyCalendar == null) { throw new System.ArgumentException("", ""); }
                    //if ((DateTime?)M3CompanyCalendar.checkDate == null) { throw new System.ArgumentException("", "No check date found"); } // difficult to detect null or missing check date, which is rare
                    LogMessage(classMsgKeyTag, "Calendar loaded");
                    this.MessageToPublishEvent("Check Date", M3CompanyCalendar.checkDate.GetDateTimeFormats()[1].ToString());
                    this.MessageToPublishEvent("Status", "Company calendar loaded");
                }
                catch (Exception e)
                {
                    LogMessage(classMsgKeyTag, "Could not load the company calendar");
                    LogMessage(MPIMsgKeyTag, e.Message);
                    this.MessageToPublishEvent("Check Date", "n/a");
                    this.MessageToPublishEvent("Status", "Could not load the company calendar - there might be no company or user session");
                    this.MessageToPublishEvent("Status", e.Message);
                    this.M3ClassCancellationNotice.Cancel();
                }
            });
            await LoadCompanyCalendarAsyncTask;
        }

        /// <summary>
        /// Loads the calendar of the company specified, resetting the company if necessary</summary>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        public async Task LoadCompanyCalendarAsync(string M3CompanyCode)
        {
            string classMsgKeyTag = "LoadCompanyCalendarAsync";
            string MPIMsgKeyTag = "MPILoadCompanyCalendar";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
            if (M3Company == null)
            {
                LogMessage(classMsgKeyTag, "Could not load the company calendar");
                LogMessage(MPIMsgKeyTag, "n/a");
                this.MessageToPublishEvent("Check Date", "n/a");
                this.MessageToPublishEvent("Status", "Could not load the company calendar - there might be no company or user session");
                this.M3ClassCancellationNotice.Cancel();
                return;
            }

            if (M3Company.co == null || M3Company.co != M3CompanyCode)
            {
                SetCompanyAsync(M3CompanyCode).Wait();
            }
            await LoadCompanyCalendarAsync();
        }

        /// <summary>
        /// Connects to the specific database, attempts to login the user, loads the company and attempts to load the calendar - no prior dependency except class instantiation</summary>
        /// <param name="dsn"> The data source name is the target M3 database name.</param>
        /// <param name="M3UserName"> An M3 username</param>
        /// <param name="M3Password"> The M3 password</param>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        public async Task LoadCompanyCalendarAsync(string dsn, string M3UserName, string M3Password, string M3CompanyCode)
        {
            SetCompanyAsync(dsn, M3UserName, M3Password, M3CompanyCode).Wait();
            await LoadCompanyCalendarAsync();
        }

        // *** AGGREGATE BATCH TOTALS ***

        /// <summary>
        /// Aggregates open batch totals (hours and amounts) for each batch and each earnings code found</summary>
        public async Task AggregateBatchTotalsAsync()
        {
            string classMsgKeyTag = "AggregateBatchTotalsAsync";
            string MPIMsgKeyTag = "MPICompanyCalendarDetail";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }

            AggregateBatchTotalsAsyncTask = Task.Factory.StartNew(() => // if the lambda called a task then use async () =>  async would otherwise still runs synchronously
            {
                try
                {
                    if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
                    this.MessageToPublishEvent("Status", "Aggregating batch totals...");
                    foreach (MILLCOMPANYLib.CCalendarDetail M3CompanyCalendarDetail in M3CompanyCalendar.Details)   //for each batch
                    {
                        foreach (MILLCOMPANYLib.CBatchTotals M3BatchTotals in M3CompanyCalendarDetail.Totals)        //for each detCode
                        {
                            allBatchTotalHours += M3BatchTotals.actualHours;
                            allBatchTotalAmounts += M3BatchTotals.actualAmount;
                            this.MessageToPublishEvent("Amount", allBatchTotalAmounts.ToString("#,##0.00"));
                            this.MessageToPublishEvent("Hours", allBatchTotalHours.ToString("#,##0.00"));
                        }
                    }
                    this.MessageToPublishEvent("Amount", allBatchTotalAmounts.ToString("#,##0.00"));
                    this.MessageToPublishEvent("Hours", allBatchTotalHours.ToString("#,##0.00"));
                    LogMessage(classMsgKeyTag, "Aggregation of batch totals complete");
                    this.MessageToPublishEvent("Status", "Aggregation of batch totals complete");
                }
                catch (Exception e)
                {
                    LogMessage(classMsgKeyTag, "Could not aggregated batch totals");
                    LogMessage(MPIMsgKeyTag, e.Message);
                    this.MessageToPublishEvent("Amount", "n/a");
                    this.MessageToPublishEvent("Hours", "n/a");
                    this.MessageToPublishEvent("Status", "Could not aggregated batch totals");
                    this.MessageToPublishEvent("Status", e.Message);
                    this.M3ClassCancellationNotice.Cancel();
                }
            });
            await AggregateBatchTotalsAsyncTask;
        }

        /// <summary>
        /// Aggregates open batch totals (hours and amounts) for the company specified, resetting the company if necessary</summary>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        public async Task AggregateBatchTotalsAsync(string M3CompanyCode)
        {
            string classMsgKeyTag = "AggregateBatchTotalsAsync";
            string MPIMsgKeyTag = "MPICompanyCalendarDetail";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }

            if (M3Company == null) // company object vs company code name below
            {
                LogMessage(classMsgKeyTag, "Could not aggregated batch totals");
                LogMessage(MPIMsgKeyTag, "n/a");
                this.MessageToPublishEvent("Amount", "n/a");
                this.MessageToPublishEvent("Hours", "n/a");
                this.MessageToPublishEvent("Status", "Could not aggregated batch totals");
                this.M3ClassCancellationNotice.Cancel();
                return;
            }

            if (M3Company.co == null || M3Company.co != M3CompanyCode)
            {
                SetCompanyAsync(M3CompanyCode).Wait();
                LoadCompanyCalendarAsync().Wait();
            }
            await AggregateBatchTotalsAsync();
        }

        /// <summary>
        /// Connects to the specific database, attempts to login the user, loads the company, loads the calendar and attempts to aggregate open batch totals (hours and amounts) - no prior dependency except class instantiation</summary>
        /// <param name="dsn"> The data source name is the target M3 database name.</param>
        /// <param name="M3UserName"> An M3 username</param>
        /// <param name="M3Password"> The M3 password</param>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        public async Task AggregateBatchTotalsAsync(string dsn, string M3UserName, string M3Password, string M3CompanyCode)
        {
            SetCompanyAsync(dsn, M3UserName, M3Password, M3CompanyCode).Wait();
            LoadCompanyCalendarAsync().Wait();
            await AggregateBatchTotalsAsync();
        }

        // *** pre-calcs payroll ***

        /// <summary>
        /// Loads the company calendar, aggregates open batch totals and pre-calculates all paychecks</summary>
        public async Task PreCalcAsync()
        {
            string classMsgKeyTag = "PreCalcAsync";
            string MPIMsgKeyTag = "MPIPaycheckCount";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }

            LoadCompanyCalendarAsync().Wait();
            AggregateBatchTotalsAsync().Wait();

            PreCalcAsyncTask = Task.Factory.StartNew(() => // if the lambda called a task then use async () =>  async would otherwise still runs synchronously
            {
                if (allBatchTotalHours == 0 && allBatchTotalAmounts == 0)
                {
                    this.MessageToPublishEvent("Check Count", "No pay found");
                    LogMessage(classMsgKeyTag, "No pay found");
                    this.MessageToPublishEvent("Status", "No pay found");
                    this.M3ClassCancellationNotice.Cancel();
                    return;
                }

                try
                {
                    if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
                    this.MessageToPublishEvent("Status", "Pre-calculating paychecks...");
                    M3CompanyCalendar.ResetPaycheckCount();
                    while (M3CompanyCalendar.PaycheckCount[""] <= 0)  //indexed by batch name, where empty string gets all members - returns a negative check count while working
                    {
                        this.MessageToPublishEvent("Check Count", (Math.Abs(M3CompanyCalendar.PaycheckCount[""])).ToString("#,##0"));
                    }
                    preCalcCheckCount = M3CompanyCalendar.PaycheckCount[""];
                    this.MessageToPublishEvent("Check Count", preCalcCheckCount.ToString("#,##0"));
                    LogMessage(classMsgKeyTag, "Pre-calc complete");
                    this.MessageToPublishEvent("Status", "Pre-calc complete");
                }
                catch (Exception e)
                {
                    LogMessage(classMsgKeyTag, "Pre-calc error");
                    LogMessage(MPIMsgKeyTag, e.Message);
                    this.MessageToPublishEvent("Check Count", "n/a");
                    this.MessageToPublishEvent("Status", "Pre-calc error...");
                    this.MessageToPublishEvent("Status", e.Message);
                    this.M3ClassCancellationNotice.Cancel();
                }
            });
            //LoadCompanyCalendarAsync().Wait();
            //AggregateBatchTotalsAsync().Wait();
            await PreCalcAsyncTask; // <-- This task is apparently fired with assignment above, because the dependencies must be called before the assignment, not prior to the await.
            //Continuewith interferes with callbacks
            //await SetCompanyAsync(M3CompanyCode).ContinueWith(t => LoadCompanyCalendarAsync()).ContinueWith(t => AggregateBatchTotalsAsync()); // interferes with callbacks (skips thems)
        }

        /// <summary>
        /// Pre-calculates all paychecks for the company specified, resetting the company if necessary (See the 1st overload for the PreCalcAsync description)</summary>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        public async Task PreCalcAsync(string M3CompanyCode)
        {
            string classMsgKeyTag = "PreCalcAsync";
            string MPIMsgKeyTag = "MPIPaycheckCount";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }

            if (M3Company == null) // company object vs company code name below
            {
                LogMessage(classMsgKeyTag, "Could not pre-calculate paychecks");
                LogMessage(MPIMsgKeyTag, "n/a");
                this.MessageToPublishEvent("Check Count", "n/a");
                this.MessageToPublishEvent("Status", "Could not pre-calculate paychecks");
                this.M3ClassCancellationNotice.Cancel();
                return;
            }

            if (M3Company.co == null || M3Company.co != M3CompanyCode)
            {
                SetCompanyAsync(M3CompanyCode).Wait();
            }
            await PreCalcAsync();
        }

        /// <summary>
        /// Connects to the specific database, attempts to login the user, loads the company, loads the calendar, aggregate batch totals and pre-calcs payroll - no prior dependency except class instantiation</summary>
        /// <param name="dsn"> The data source name is the target M3 database name.</param>
        /// <param name="M3UserName"> An M3 username</param>
        /// <param name="M3Password"> The M3 password</param>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        public async Task PreCalcAsync(string dsn, string M3UserName, string M3Password, string M3CompanyCode)
        {
            SetCompanyAsync(dsn, M3UserName, M3Password, M3CompanyCode).Wait();
            await PreCalcAsync();
        }

        // *** RUN REPORTS FROM TREE ***

        /// <summary>
        /// Runs reports existing in an M3 tree folder using a path override option and a partial-text label or name filter</summary>
        /// <param name="M3CompanyTreePath"> M3 company-level tree path, such as @"/Reporting/Preprocess Reports"; The M3 class prefixes "/UIRoot/CompanySets/Companies/") </param>
        /// <param name="TargetOutputUNCPath"> UNC path that overrides any existing report output folder specification, such as @"\\HOST3\Millennium Shared Files\Temp" (blank = existing path)</param>
        /// <param name="M3ReportLabelOrTitlePattern"> This is a pattern-matching filter to select specific reports by looking for text contained within the label or name (blank = no filter, running all reports within the folder)</param>
        public async Task RunReportsFromTreeAsync(string M3CompanyTreePath, string TargetOutputUNCPath, string M3ReportLabelOrTitlePattern) //treePath As String, uncPath As String, label As String
        {
            string classMsgKeyTag = "RunReportsFromTreeAsync";
            string MPIMsgKeyTag = "MPIRunReport";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
            int reportsCount = 0;

            ReportJobParametersStruct TricoreReportJob = ReturnReportJobParametersStruct(M3Company.co, M3CompanyTreePath, TargetOutputUNCPath, M3ReportLabelOrTitlePattern);
            //MILLREPORTLib.RReport M3Report;

            try
            {
                M3TreeHelper = (MILLTREELib.TreeHelper)M3System.TreeHelper; // uses the system tree with login access
                M3FolderNode = M3TreeHelper.GetContextNodeFromPath(TricoreReportJob.M3CompanyTreePath, M3Company.guidColumn);
                // MILLTREELib.IMPICollection get_ContextContents(MILLTREELib.IMPIEntity pCo) - It seems possible access this same node structure (if found) within other companies
                M3FolderNodeItemsCollection = M3FolderNode.ContextContents[(MILLTREELib.IMPIEntity)M3Company]; // collect the items in the folder -- M3FolderNode.Parent ?
            }
            catch (Exception e)
            {
                LogMessage(classMsgKeyTag, "Could not find the M3 tree path: " + TricoreReportJob.M3CompanyTreePath);
                LogMessage(MPIMsgKeyTag, e.Message);
                this.MessageToPublishEvent("Status", "Could not find the M3 tree path: " + TricoreReportJob.M3CompanyTreePath);
                this.MessageToPublishEvent("Status", e.Message);
                this.M3ClassCancellationNotice.Cancel();
                return;
            }

            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
            this.MessageToPublishEvent("Status", "Running reports...");

            // ******************************

            ////This was a neat approach but it fire reports before Millennium can process them...
            //List<Task> RunReportsTaskList = new List<Task>();
            //Task RunReportsTask;

            //foreach (MILLTREELib.TNode M3Node in M3FolderNodeItemsCollection)
            //{
            //    if (M3Node.NodeType.label == "Report"
            //        && (TricoreReportJob.M3ReportLabelOrTitlePattern == ""
            //            || M3Node.label.Contains(TricoreReportJob.M3ReportLabelOrTitlePattern)
            //            || M3Node.description.Contains(TricoreReportJob.M3ReportLabelOrTitlePattern)
            //            )
            //        )
            //    {
            //        M3Report = (MILLREPORTLib.RReport)M3Node.AsReport;
            //        if (TricoreReportJob.TargetOutputUNCPath != "")
            //        {
            //            if (M3Report.PhysicalReport.label.Substring(0, 5) == "TRI_x") // Tricore specific custom reporting system
            //            {
            //                M3Report.Properties.addlFormValueByName["YYYUNCPath"] = TricoreReportJob.TargetOutputUNCPath;
            //            }
            //            else
            //            {
            //                M3Report.Properties.outputDirSpecific = TricoreReportJob.TargetOutputUNCPath;
            //            }
            //        }

            //        RunReportsTask = Task.Factory.StartNew(() =>
            //        {
            //            try
            //            {
            //                M3Report.RunReport((MILLREPORTLib.IMPIEntity)M3Company, "NoPrint"); // Tricore specific no-print printer
            //                reportsCount++;
            //                this.MessageToPublishEvent("Reports Count", reportsCount.ToString("N0", ci));
            //            }
            //            catch (Exception e)
            //            {
            //                this.MessageToPublishEvent("Status", "Could not run " + M3Report.AsNode.DescFieldValue);
            //                this.MessageToPublishEvent("Status", e.Message);
            //            }

            //        });
            //        RunReportsTaskList.Add(RunReportsTask); // <-- this fires reports before Millennium can process them - some don't run
            //    }
            //}

            //RunReportsFromTreeAsyncTask = Task.WhenAll(RunReportsTaskList);

            // ******************************

            RunReportsFromTreeAsyncTask = Task.Factory.StartNew(() =>
            {
                foreach (MILLTREELib.TNode M3Node in M3FolderNodeItemsCollection)
                {
                    if (M3ClassCancellationNotice.IsCancellationRequested) { return; }
                    if (M3Node.NodeType.label == "Report"
                        && (TricoreReportJob.M3ReportLabelOrTitlePattern == ""
                            || M3Node.label.Contains(TricoreReportJob.M3ReportLabelOrTitlePattern)
                            || M3Node.description.Contains(TricoreReportJob.M3ReportLabelOrTitlePattern)
                            )
                        )
                    {
                        M3Report = (MILLREPORTLib.RReport)M3Node.AsReport;
                        if (TricoreReportJob.TargetOutputUNCPath != "")
                        {
                            if (M3Report.PhysicalReport.label.Substring(0, 5) == "TRI_x") // Tricore specific custom reporting system
                            {
                                M3Report.Properties.addlFormValueByName["YYYUNCPath"] = TricoreReportJob.TargetOutputUNCPath;
                            }
                            else
                            {
                                M3Report.Properties.outputDirSpecific = TricoreReportJob.TargetOutputUNCPath;
                            }
                            //M3Report.Properties.outputFormat = "PDF";
                            //M3Report.Properties.outputDirFlag = 1;
                            //M3Report.Properties.outputDirSpecific = TricoreReportJob.TargetOutputUNCPath;
                            //M3Report.Properties.outputFilenameFlag = 1;
                            //M3Report.Properties.outputFilenameSpecific = "[COMPANY.CO] [REPORT.CHECKDATESTART(%m-%d-%y)].pdf";
                            //NOT NECESSARY TO SAVE TO RUN
                            //M3Report.Properties.SaveChanges(); // not necessary to fire this way here
                        }

                        try
                        {
                            // Millennium seems to prefer a synchronous type use of RunReport vs something like RunReportsTaskList.Add(RunReportsTask), which can use WhenAll
                            M3Report.RunReport((MILLREPORTLib.IMPIEntity)M3Company, "NoPrint"); // Tricore specific no-print printer
                            reportsCount++;
                            this.MessageToPublishEvent("Reports Count", reportsCount.ToString("N0", ci));
                        }
                        catch (Exception e)
                        {
                            this.MessageToPublishEvent("Status", "Could not run " + M3Report.AsNode.DescFieldValue);
                            this.MessageToPublishEvent("Status", e.Message);
                        }
                    }
                }
            });

            //this works to pickup the end of the task
            await RunReportsFromTreeAsyncTask.ContinueWith(t => Task.Factory.StartNew(() =>
            {
                LogMessage(classMsgKeyTag, reportsCount.ToString("N0", ci) + " report" + (reportsCount == 1 ? "" : "s") + " queued");
                this.MessageToPublishEvent("Status", reportsCount.ToString("N0", ci) + " report" + (reportsCount == 1 ? "" : "s") + " queued");
            }));
        }

        // ******************************

        /// <summary>
        /// Runs reports existing in an M3 tree folder for the company specified, resetting the company if necessary (See the 1st overload for the RunReportsFromTreeAsync description)</summary>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        /// <param name="M3CompanyTreePath"> M3 company-level tree path, such as @"/Reporting/Preprocess Reports"; The M3 class prefixes "/UIRoot/CompanySets/Companies/") </param>
        /// <param name="TargetOutputUNCPath"> UNC path that overrides any existing report output folder specification, such as @"\\HOST3\Millennium Shared Files\Temp" (blank = existing path)</param>
        /// <param name="M3ReportLabelOrTitlePattern"> This is a pattern-matching filter to select specific reports by looking for text contained within the label or name (blank = no filter, running all reports within the folder)</param>
        public async Task RunReportsFromTreeAsync(string M3CompanyCode, string M3CompanyTreePath, string TargetOutputUNCPath, string M3ReportLabelOrTitlePattern) //treePath As String, uncPath As String, label As String
        {
            string classMsgKeyTag = "RunReportsFromTreeAsync";
            string MPIMsgKeyTag = "MPIRunReport";
            if (M3ClassCancellationNotice.IsCancellationRequested) { return; }

            if (M3Company == null) // company object vs company code name below -- this probably cannot happen - a null paramter might get flagged by VS
            {
                LogMessage(classMsgKeyTag, "Running reports for a specific company requires a company code. Please provide a company code.");
                LogMessage(MPIMsgKeyTag, "n/a");
                this.MessageToPublishEvent("Report Count", "0");
                this.MessageToPublishEvent("Status", "Could not run reports, because the company parameter was null");
                this.M3ClassCancellationNotice.Cancel();
                return;
            }

            if (M3Company.co == null || M3Company.co != M3CompanyCode)
            {
                SetCompanyAsync(M3CompanyCode).Wait(); //blocking but must be complete before running rpeorts
            }
            await RunReportsFromTreeAsync(M3CompanyTreePath, TargetOutputUNCPath, M3ReportLabelOrTitlePattern);
        }

        /// <summary>
        /// Connects to the specific database, attempts to login the user, loads the company, and attempts to run reports existing in an M3 tree folder - no prior dependency except class instantiation</summary>
        /// <param name="dsn"> The data source name is the target M3 database name.</param>
        /// <param name="M3UserName"> An M3 username</param>
        /// <param name="M3Password"> The M3 password</param>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        /// <param name="M3CompanyTreePath"> M3 company-level tree path, such as @"/Reporting/Preprocess Reports"; The M3 class prefixes "/UIRoot/CompanySets/Companies/") </param>
        /// <param name="TargetOutputUNCPath"> UNC path that overrides any existing report output folder specification, such as @"\\HOST3\Millennium Shared Files\Temp" (blank = existing path)</param>
        /// <param name="M3ReportLabelOrTitlePattern"> This is a pattern-matching filter to select specific reports by looking for text contained within the label or name (blank = no filter, running all reports within the folder)</param>
        public async Task RunReportsFromTreeAsync(string dsn, string M3UserName, string M3Password, string M3CompanyCode, string M3CompanyTreePath, string TargetOutputUNCPath, string M3ReportLabelOrTitlePattern) //treePath As String, uncPath As String, label As String
        {
            SetCompanyAsync(dsn, M3UserName, M3Password, M3CompanyCode).Wait(); //blocking but must be complete before running rpeorts
            await RunReportsFromTreeAsync(M3CompanyTreePath, TargetOutputUNCPath, M3ReportLabelOrTitlePattern);
        }


        /// <summary>
        /// Pre-calculates all paychecks and runs reports existing in an M3 tree folder using a path override option and a partial-text label or name filter (treefolder, outputUNCPath, pattern);</summary>
        /// <param name="M3CompanyTreePath"> M3 company-level tree path, such as @"/Reporting/Preprocess Reports"; The M3 class prefixes "/UIRoot/CompanySets/Companies/") </param>
        /// <param name="TargetOutputUNCPath"> UNC path that overrides any existing report output folder specification, such as @"\\HOST3\Millennium Shared Files\Temp" (blank = existing path)</param>
        /// <param name="M3ReportLabelOrTitlePattern"> This is a pattern-matching filter to select specific reports by looking for text contained within the label or name (blank = no filter, running all reports within the folder)</param>
        public async Task RunReportsFromTreeWithPreCalcAsync(string M3CompanyTreePath, string TargetOutputUNCPath, string M3ReportLabelOrTitlePattern) //treePath As String, uncPath As String, label As String
        {
            PreCalcAsync().Wait();
            await RunReportsFromTreeAsync(M3CompanyTreePath, TargetOutputUNCPath, M3ReportLabelOrTitlePattern);
        }

        /// <summary>
        /// Runs reports existing in an M3 tree folder for the company specified, resetting the company if necessary (See the 1st overload for the RunReportsFromTreeAsync description)</summary>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        /// <param name="M3CompanyTreePath"> M3 company-level tree path, such as @"/Reporting/Preprocess Reports"; The M3 class prefixes "/UIRoot/CompanySets/Companies/") </param>
        /// <param name="TargetOutputUNCPath"> UNC path that overrides any existing report output folder specification, such as @"\\HOST3\Millennium Shared Files\Temp" (blank = existing path)</param>
        /// <param name="M3ReportLabelOrTitlePattern"> This is a pattern-matching filter to select specific reports by looking for text contained within the label or name (blank = no filter, running all reports within the folder)</param>
        public async Task RunReportsFromTreeWithPreCalcAsync(string M3CompanyCode, string M3CompanyTreePath, string TargetOutputUNCPath, string M3ReportLabelOrTitlePattern) //treePath As String, uncPath As String, label As String
        {
            PreCalcAsync(M3CompanyCode).Wait();
            await RunReportsFromTreeAsync(M3CompanyTreePath, TargetOutputUNCPath, M3ReportLabelOrTitlePattern);
        }

        /// <summary>
        /// Connects to the specific database, attempts to login the user, loads the company, loads the calendar, aggregates batch totals, pre-calcs payroll and attempts to run reports existing in an M3 tree folder - no prior dependency except class instantiation</summary>
        /// <param name="dsn"> The data source name is the target M3 database name.</param>
        /// <param name="M3UserName"> An M3 username</param>
        /// <param name="M3Password"> The M3 password</param>
        /// <param name="M3CompanyCode"> An M3 company code (case sensitive)</param>
        /// <param name="M3CompanyTreePath"> M3 company-level tree path, such as @"/Reporting/Preprocess Reports"; The M3 class prefixes "/UIRoot/CompanySets/Companies/") </param>
        /// <param name="TargetOutputUNCPath"> UNC path that overrides any existing report output folder specification, such as @"\\HOST3\Millennium Shared Files\Temp" (blank = existing path)</param>
        /// <param name="M3ReportLabelOrTitlePattern"> This is a pattern-matching filter to select specific reports by looking for text contained within the label or name (blank = no filter, running all reports within the folder)</param>
        public async Task RunReportsFromTreeWithPreCalcAsync(string dsn, string M3UserName, string M3Password, string M3CompanyCode, string M3CompanyTreePath, string TargetOutputUNCPath, string M3ReportLabelOrTitlePattern) //treePath As String, uncPath As String, label As String
        {
            PreCalcAsync(dsn, M3UserName, M3Password, M3CompanyCode).Wait();
            await RunReportsFromTreeAsync(M3CompanyTreePath, TargetOutputUNCPath, M3ReportLabelOrTitlePattern);
        }
    }
}

