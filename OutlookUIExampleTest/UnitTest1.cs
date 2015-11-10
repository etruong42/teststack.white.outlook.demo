using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using TestStack.White;
using TestStack.White.UIItems.Finders;
using TestStack.White.Factory;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Threading;
using System.Linq;

namespace OutlookUIExampleTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            // launch Outlook 2010
            var outlookPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                @"Microsoft Office\Office14\OUTLOOK.EXE");
            var application = Application.Launch(outlookPath);

            // get explorer window
            var explorer = application.GetWindow(
                SearchCriteria.ByClassName("rctrl_renwnd32"),
                InitializeOption.NoCache);

            // click "New E-mail" button to start composing new email
            explorer.Get(SearchCriteria.ByText("New E-mail")).Click();

            // get composer window
            var composer = application.GetWindow(
                SearchCriteria.ByClassName("rctrl_renwnd32").AndByText("Untitled - Message (HTML) "),
                InitializeOption.NoCache);            

            // fill out "To" field
            var toField = composer.Get(SearchCriteria.ByClassName("RichEdit20WPT").AndByText("To"));
            toField.Enter("someone@example.com");

            // fill out "Subject" field
            var subjectField = composer.Get(SearchCriteria.ByClassName("RichEdit20WPT").AndByText("Subject:"));
            subjectField.Enter("automated UI email");

            // change focus to get Outlook process registered in running object table
            // https://social.msdn.microsoft.com/Forums/office/en-US/0d8f9642-50bc-4656-af32-84d62068305d/outlook-2010-and-registering-in-the-running-object-table?forum=outlookdev
            var windows = WindowFactory.Desktop.DesktopWindows();
            var desktop = windows.Last().GetElement(SearchCriteria.ByClassName("SysListView32"));
            desktop.SetFocus();

            Outlook.Application outlookCom = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;

            var sentMailItem = outlookCom.ActiveInspector().CurrentItem as Outlook.MailItem;
            var body = sentMailItem.HTMLBody;
            var index = body.IndexOf(@"</body", StringComparison.InvariantCultureIgnoreCase);
            sentMailItem.HTMLBody = body.Insert(index, "this email was sent via automated UI");


            composer.Get(SearchCriteria.ByText("Send").AndByClassName("Button")).Click();

            // give Outlook time to send off the email
            Thread.Sleep(TimeSpan.FromSeconds(5));
            
            application.WaitWhileBusy();
            explorer.Close();
        }
    }
}
