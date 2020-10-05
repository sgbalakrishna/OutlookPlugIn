using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

namespace mailBoxWizard
{
    class Program
    {
        public static object finaldata { get; private set; }
        public string EmailSender { get; private set; }
        public double freq { get; private set; }
        public int Counts { get; private set; }


        public static List<Program> PrepareData()
        {
            var mails = OutlookEmails.ReadMailItems();
            List<string> senders = new List<string>();

            List<string> domains = new List<string>();
            List<int> avgDurations = new List<int>();

            List<Program> listEmailDetailsfinal = new List<Program>();
            Program emailDetailsfinal;


            foreach (var mail in mails)
            {
                senders.Add(mail.EmailFrom.ToString());
            }

            IEnumerable<string> unqSenders = senders.Distinct();


            foreach (string sender in unqSenders)
            {
                int cnt = 0;
                emailDetailsfinal = new Program();
                List<DateTime> dt = new List<DateTime>();
                List<int> durations = new List<int>();
                durations.Add(0);

                foreach (var ml in mails)
                {
                    if (ml.EmailFrom == sender)
                    {
                        dt.Add(ml.RecievedOn);
                        cnt += 1;
                    }
                }

                dt = SortAscending(dt);

                for (int a = 0; a < (dt.Count) - 1; a++)
                {
                    TimeSpan x = dt[a + 1].Subtract(dt[a]);
                    var duration = Convert.ToInt32(x.TotalMinutes / (1440));
                    durations.Add(duration);

                }

                emailDetailsfinal.EmailSender = sender;
                emailDetailsfinal.Counts = cnt;
                emailDetailsfinal.freq = Math.Round(durations.Average(), 0);
                listEmailDetailsfinal.Add(emailDetailsfinal);

            }

            return listEmailDetailsfinal;
        }

        private static List<DateTime> SortAscending(List<DateTime> li)
        {
            li.Sort((a, b) => a.CompareTo(b));
            return li;
        }
    }

}

