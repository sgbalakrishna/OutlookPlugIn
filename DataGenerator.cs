using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Outlook.Application;

namespace mailBoxWizard
{
    public class OutlookEmails
    {
        public string EmailFrom { get; set; } = string.Empty;
        public string EmailSubject { get; set; } = string.Empty;
        public dynamic RecievedOn { get; set; }
        public string EmailBody { get; set; } = string.Empty;
        public string Sender { get; internal set; }
        public double freq { get; internal set; }

        public static List<OutlookEmails> ReadMailItems()

        {
            Application outlookapplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;

            Items mailItems = null;
            List<OutlookEmails> listEmailDetails = new List<OutlookEmails>();
            OutlookEmails emailDetails;

            try
            {
                outlookapplication = new Application();
                outlookNamespace = outlookapplication.GetNamespace("MAPI");

                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                mailItems = inboxFolder.Items;
                ProgressBar pBar1 = new ProgressBar();
                pBar1.Minimum = 1;

                pBar1.Maximum = mailItems.Count;
                pBar1.Value = 1;
                pBar1.Step = 1;
                
                
                Form1 pbarForm = new Form1();
                pBar1.Width = 677;
                
                pbarForm.Controls.Add(pBar1);
                pbarForm.StartPosition = FormStartPosition.CenterScreen;
                pbarForm.Show();


                for (int j = 1; j < mailItems.Count; j++)

                {
                    emailDetails = new OutlookEmails();
                    emailDetails.EmailSubject = mailItems[j].Subject;
                    emailDetails.RecievedOn = mailItems[j].ReceivedTime;
                    if (mailItems[j].senderEmailType == "EX")
                    {
                        emailDetails.EmailFrom = "valyue.de";
                    }
                    if (mailItems[j].senderEmailType == "SMTP")
                    {
                        string after = "@";
                        string x = mailItems[j].SenderEmailAddress;
                        string final = x.Substring(x.LastIndexOf(after) + 1);
                        emailDetails.EmailFrom = final.ToLower();
                    }
                    listEmailDetails.Add(emailDetails);
                    pBar1.PerformStep();

                }
                System.Threading.Thread.Sleep(2);
                pbarForm.Close();

            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookapplication);
            }
            return listEmailDetails;
        }

        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;

            }
        }
    }
}
