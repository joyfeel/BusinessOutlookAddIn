using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;


namespace BusinessOutlookAddIn
{
    static class Constants
    {
        public const int AttachmentContentLength = 2048;             // just in case
        public const string EncryptionHeader = "OSR__DS_FILE_HDR";

        public const string WarningMessagePromptTitle = "附件提醒";
        public const string WarningMessagePromptContent = "仍要傳送信件嗎?";

        public const string WarningMessagePromptNotMatchRecipients = "附檔的專案可能不屬於此收件者";
        public const string WarningMessagePromptUndecrypted = "附件尚未解密";
        public const string WarningMessagePromptForgetAttachment = "可能忘記附加檔案";  // TODO
        public const string WarningMessagePromptFormatIssue = "附檔可能是未翻譯的PPT或檔名命名規則錯誤";
    }

    public class ListMap<T, V> : List<KeyValuePair<T, V>>
    {
        public void Add(T key, V value)
        {
            Add(new KeyValuePair<T, V>(key, value));
        }

        public List<V> Get(T key)
        {
            return FindAll(p => p.Key.Equals(key)).ConvertAll(p => p.Value);
        }
    }

    public partial class ThisAddIn
    {
        private Outlook.Inspectors inspectors;

        private string RemoveSpecialCharacters(string str)
        {
            StringBuilder sb = new StringBuilder();

            foreach (char c in str)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_')
                {
                    sb.Append(c);
                }
            }

            return sb.ToString();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private bool isNotMatchRecipients(Outlook.Attachment attachment, Outlook.MailItem mailItem)
        {
            string fileName = attachment.FileName;                                      // N0861_Hairpin_Chip_R2_Q_0409
            string currentYear = DateTime.Now.Year.ToString();                          // 2020
            string newFileName = fileName[0] + currentYear + fileName.Substring(1, 4);  // N20200861

            const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            Outlook.Recipients recips = mailItem.Recipients;

            var dummyDB = new ListMap<string, string> {
                { "N20200861", "unihancorp.com" },
                { "N20200861", "pegatroncorp.com" },
                { "N20200861", "hotmail.com" },
                { "N20200862", "gmail.com" },
                { "N20200863", "gmail.com" },
            };

            bool hasError = false;
            bool innerLoopError = true;
            foreach (Outlook.Recipient recip in recips)
            {
                string recipMail = recip.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS).ToString();
                string recipDomain = recipMail.Split('@')[1];

                foreach (var pair in dummyDB)
                {
                    string ProjectNumber = pair.Key;
                    string Domain = pair.Value;

                    if (ProjectNumber == newFileName)
                    {
                        innerLoopError &= (Domain != recipDomain);
                    }
                }

                hasError |= innerLoopError;
            }

            return hasError;
        }

        private bool isUndecryptedAttachment(Outlook.Attachment attachment, Outlook.MailItem mailItem)
        {
            const string PR_ATTACH_DATA_BIN = "http://schemas.microsoft.com/mapi/proptag/0x37010102";

            // Retrieve the attachment as an array of bytes.
            byte[] attachmentData = attachment.PropertyAccessor.GetProperty(PR_ATTACH_DATA_BIN);

            int attachmentCount = Constants.AttachmentContentLength;

            if (attachment.Size < attachmentCount)
            {
                attachmentCount = attachment.Size;
            }

            string attachmentContent = RemoveSpecialCharacters(Encoding.UTF8.GetString(attachmentData, 0, attachmentCount));

            bool isMatch = System.Text.RegularExpressions.Regex.IsMatch(attachmentContent, Constants.EncryptionHeader);

            return isMatch;
        }

        void Application_ItemSend(object Item, ref bool Cancel)
        {
            Outlook.MailItem mailItem = Item as Outlook.MailItem;

            if (mailItem != null)
            {
                var attachments = mailItem.Attachments; ;
                string message = "";
                bool hasNotMatchRecipientsError = false;
                bool hasUndecryptedError = false;

                if (attachments == null)
                {
                    return;
                }

                foreach (Outlook.Attachment attachment in attachments)
                {
                    if (isNotMatchRecipients(attachment, mailItem) == true)
                    {
                        if (!hasNotMatchRecipientsError)
                        {
                            message += Constants.WarningMessagePromptNotMatchRecipients + "\n";
                        }

                        hasNotMatchRecipientsError = true;
                    }

                    if (isUndecryptedAttachment(attachment, mailItem) == true)
                    {
                        if (!hasUndecryptedError)
                        {
                            message += Constants.WarningMessagePromptUndecrypted + "\n";
                        }

                        hasUndecryptedError = true;
                    }
                }

                bool hasError = hasUndecryptedError || hasNotMatchRecipientsError;

                if (hasError)
                {
                    message += Constants.WarningMessagePromptContent;

                    DialogResult result = MessageBox.Show(
                        message,
                        Constants.WarningMessagePromptTitle,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (result == DialogResult.No)
                    {
                        Cancel = true;
                    }
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
