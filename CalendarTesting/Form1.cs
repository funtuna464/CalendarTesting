using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Util;
using System.IO;
using System.Threading;

namespace WindowsFormsApp1
{
    public partial class frm1 : Form
    {
        static string[] Scopes = { CalendarService.Scope.CalendarEvents };
        static string ApplicationName = "Semester Class Adding Script";

        public frm1()
        {
            InitializeComponent();
        }

        private void frm1_Load(object sender, EventArgs e)
        {

        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            btnStart.Text = "Working...";
            string title = txtName.Text;
            string locat = txtLoc.Text;
            DateTime startTime = datClassStart.Value;
            DateTime endTime = DatClassEnd.Value;
            DateTime startSem = datSemStart.Value;
            DateTime endSem = datSemEnd.Value;
            bool mon = chkMon.Checked;
            bool tue = chkTue.Checked;
            bool wed = chkWed.Checked;
            bool thu = chkThu.Checked;
            bool fri = chkFri.Checked;
            //check if any of the fields are blank
            if (title.Length == 0 || locat.Length == 0 || !(mon || tue || wed || thu || fri))
            {
                btnStart.Text = "One or more fields are blank";
                return;
            }
            //nudge the startSem to match the start day
            switch (startSem.DayOfWeek.ToString())
            {
                case "Monday":
                    if (mon) { }
                    else if (tue) { startSem = startSem.AddDays(1); }
                    else if (wed) { startSem = startSem.AddDays(2); }
                    else if (thu) { startSem = startSem.AddDays(3); }
                    else if (fri) { startSem = startSem.AddDays(4); }
                    break;
                case "Tuesday":
                    if (mon) { startSem = startSem.AddDays(6); }
                    else if (tue) { }
                    else if (wed) { startSem = startSem.AddDays(1); }
                    else if (thu) { startSem = startSem.AddDays(2); }
                    else if (fri) { startSem = startSem.AddDays(3); }
                    break;
                case "Wednesday":
                    if (mon) { startSem = startSem.AddDays(5); }
                    else if (tue) { startSem = startSem.AddDays(6); }
                    else if (wed) { }
                    else if (thu) { startSem = startSem.AddDays(1); }
                    else if (fri) { startSem = startSem.AddDays(2); }
                    break;
                case "Thursday":
                    if (mon) { startSem = startSem.AddDays(4); }
                    else if (tue) { startSem = startSem.AddDays(5); }
                    else if (wed) { startSem = startSem.AddDays(6); }
                    else if (thu) { }
                    else if (fri) { startSem = startSem.AddDays(1); }
                    break;
                case "Friday":
                    if (mon) { startSem = startSem.AddDays(3); }
                    else if (tue) { startSem = startSem.AddDays(4); }
                    else if (wed) { startSem = startSem.AddDays(5); }
                    else if (thu) { startSem = startSem.AddDays(6); }
                    else if (fri) { }
                    break;
            }
            //Choose the color

            int col = 1;
            
            RadioButton radioBtn = this.Controls.OfType<RadioButton>()
                                       .Where(x => x.Checked).FirstOrDefault();
            if (radioBtn != null)
            {
                switch (radioBtn.Name)
                {
                    case "rad2":
                        col = 2;
                        break;
                    case "rad3":
                        col = 3;
                        break;
                    case "rad4":
                        col = 4;
                        break;
                    case "rad5":
                        col = 5;
                        break;
                    case "rad6":
                        col = 6;
                        break;
                    case "rad7":
                        col = 7;
                        break;
                    case "rad9":
                        col = 9;
                        break;
                    case "rad10":
                        col = 10;
                        break;
                    case "rad11":
                        col = 11;
                        break;

                        //Your switch structure here ...

                }
            }
            //Create the initial start and end DateTime
            DateTime startClass = new DateTime(startSem.Year, startSem.Month, startSem.Day, startTime.Hour, startTime.Minute, 0);
            DateTime endClass = new DateTime(startSem.Year, startSem.Month, startSem.Day, endTime.Hour, endTime.Minute, 0);
            createEvent(title, locat, startClass, endClass, createRecur(endSem, mon, tue, wed, thu, fri), col);
            btnStart.Text = "Done!";
        }

        void createEvent(String tit, String loc, DateTime start, DateTime end, String recur, int color)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            Event theclass = new Event()
            {
                Summary = tit,
                Location = loc,
                Start = new EventDateTime()
                {
                    DateTime = start,
                    TimeZone = "America/New_York"
                },
                End = new EventDateTime()
                {
                    DateTime = end,
                    TimeZone = "America/New_York"
                },
                Recurrence = new String[] { recur },
                Reminders = new Event.RemindersData()
                {
                    UseDefault = false
                },
                ColorId = color.ToString()
            };
            Event recurringEvent = service.Events.Insert(theclass, "primary").Execute();
            Console.WriteLine("done?");
        }

        String createRecur(DateTime end, bool m, bool tu, bool w, bool th, bool f)
        {
            /*
             * We need to convert DateTime to an RRULE format, so we need to do a little string conversion
             * Starts out as 03/01/2008
             * Need it to be 20080301
             */
            String dat = end.ToString("yyyyMMdd");
            String days = "";
            if (m)
            {
                days += "MO,";
            }
            if (tu)
            {
                days += "TU,";
            }
            if (w)
            {
                days += "WE,";
            }
            if (th)
            {
                days += "TH,";
            }
            if (f)
            {
                days += "FR,";
            }
            days = days.TrimEnd(',');
            return "RRULE:FREQ=WEEKLY;BYDAY=" + days + ";INTERVAL=1;UNTIL=" + dat + "T050000Z";
        }

        void reset()
        {
            btnStart.Text = "Add Class";
        }

        private void txtName_TextChanged(object sender, EventArgs e)
        {
            reset();
        }

        private void txtLoc_TextChanged(object sender, EventArgs e)
        {
            reset();
        }

        private void btnRev_Click(object sender, EventArgs e)
        {
            UserCredential credential;
            string credPath = "token.json";
            if (Directory.Exists(credPath))
            {
                using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Google Calendar API service.
                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
                Directory.Delete(credPath, true);
            }
        }
    }
}
