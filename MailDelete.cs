using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Outlook.Application;

namespace mailBoxWizard
{
    public class RemoveMailItems
    {
        public string EmailFrom { get; set; } = string.Empty;

        public string EmailSubject { get; set; } = string.Empty;
        public dynamic RecievedOn { get; set; }
        public string EmailBody { get; set; } = string.Empty;
        public string Sender { get; internal set; }


        internal static void deleteMailItems(List<string> checkedItems)

        {
            Application outlookapplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;
            Items mailItems = null;


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
                pBar1.Width = 500;
                Form1 pbarForm = new Form1();
                pbarForm.Width = pBar1.Width;
                pbarForm.Height = 75;
                pbarForm.Controls.Add(pBar1);
                pbarForm.Show();

                for (int j = 1; j < mailItems.Count; j++)
                {
                    string after = "@";
                    string x = mailItems[j].SenderEmailAddress;
                    string final = x.Substring(x.LastIndexOf(after) + 1).ToString();
                    if (checkedItems.Any(w => final.Contains(w)))
                    {
                        mailItems[j].delete();
                    }
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
