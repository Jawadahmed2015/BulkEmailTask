using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Pdf;
using Aspose.Pdf.Generator;
using System.Net.Mail;
using System.Net;

namespace Gridview_Export_To_CSV
{
    public partial class Gridview_To_CSV : System.Web.UI.Page
    {
        /// <summary>
        /// //////// global variables so that we dont have to make extra varibles ///////////////////////////
        /// </summary>

        DataTable dt = new DataTable();
        string fileName = string.Empty;
        string FilePath = string.Empty;
        string[] filenames = { };

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                setPlaceHold();
                Session["listUsers"] = null;
                Session["listEmails"] = null;
            }
            else
                message.Visible = false;
        }

        /// <summary>
        /// event  when user click on upload button to upload data from required csv file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        protected void btnContactUpload_Click(object sender, EventArgs e)
        {
            //check to make sure a file is selected
            if (contactsFileUpload.HasFile)
            {
                //create the path to save the file to
                fileName = Path.Combine(Server.MapPath("~/uploads"), contactsFileUpload.FileName);
                //save the file to our local path
                contactsFileUpload.SaveAs(fileName);
                filenames = Directory.GetFiles(Server.MapPath("~/uploads"));
                if (filenames.Length > 0) // checking fileUploader has file or not , or user just click on upload button
                {

                    //function which read csvfile and return data in datatable
                    dt = GetDataTabletFromCSVFile(fileName);
                    Session.Add("CsvDataTable", dt);
                    Session.Add("EmailfileName", fileName);

                    message.Text = "Contact Uploaded Successfully";
                    message.Visible = true;
                    message.Attributes.Add("class", "alert alert-success");

                }
            }
            else if (dt.Rows.Count <= 0)
            {
                LoadData();
                message.Text = "Please First Select CSV File";
                message.Visible = true;
                message.Attributes.Add("class", "alert alert-danger");

            }
        }

        /// <summary>
        /// // fill required fields when requrie 
        /// </summary>
        private void LoadData()
        {
            setPlaceHold();
            if (Session["EmailfileName"] == null)
                return;
            fileName = Session["EmailfileName"].ToString();
            dt = GetDataTabletFromCSVFile(fileName);

        }



        /// <summary>
        /// /// reading csv file using TextFieldParser and return data to datatable
        /// </summary>
        /// <param name="csv_file_path"></param>
        /// <returns></returns>
        private DataTable GetDataTabletFromCSVFile(string csv_file_path)
        {
            try
            {
                using (TextFieldParser csvFileReader = new TextFieldParser(csv_file_path))
                {
                    csvFileReader.SetDelimiters(new string[] { "," });
                    csvFileReader.HasFieldsEnclosedInQuotes = true;

                    //reading column names in the file
                    string[] requiredcolFields = csvFileReader.ReadFields();
                    foreach (string column in requiredcolFields)
                    {
                        DataColumn datecolumn = new DataColumn(column);
                        datecolumn.AllowDBNull = true;
                        dt.Columns.Add(datecolumn);
                    }
                    while (!csvFileReader.EndOfData) // read untill file is compeleted
                    {
                        string[] fieldData = csvFileReader.ReadFields();
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] == "")    // checking  empty value and set as null
                            {
                                fieldData[i] = null;
                            }
                        }
                        listUsers.Items.Add(fieldData[0]);  // adding in requirement adding values in list  
                        listEmails.Items.Add(fieldData[1]); // adding in requirement adding values in list 
                        dt.Rows.Add(fieldData);

                        Session["listUsers"] = listUsers;
                        Session["listEmails"] = listEmails;
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return dt;
        }

        /// <summary>
        /// email content uploader
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        protected void emailButtonUpload_Click(object sender, EventArgs e)
        {
            if (emailTemplateFileUpload.HasFile)
            {
                FilePath = Server.MapPath("~/uploads");
                fileName = Path.Combine(Server.MapPath("~/uploads"), emailTemplateFileUpload.FileName);
                emailTemplateFileUpload.SaveAs(fileName);
                filenames = Directory.GetFiles(Server.MapPath("~/uploads"));
                if (fileName.Length > 0)
                {
                    Aspose.Words.Document doc = new Aspose.Words.Document(FilePath + "\\" + emailTemplateFileUpload.FileName);
                    // Set an option to export form fields as plain text, not as HTML input elements.
                    Aspose.Words.Saving.HtmlSaveOptions options = new Aspose.Words.Saving.HtmlSaveOptions(Aspose.Words.SaveFormat.Html);
                    options.ExportTextInputFormFieldAsText = true;
                    doc.Save(FilePath + "\\MailMergeHtml.html", options);

                    string body = string.Empty;
                    using (System.IO.StreamReader reader = new System.IO.StreamReader(FilePath + "\\MailMergeHtml.html"))
                    {
                        body = reader.ReadToEnd();
                        mailMergeHtml.Text = body;
                    }
                    GeneratePdf(body); //generate required pdf
                }
                message.Text = "MailMerge Template Loaded successfully";
                message.Visible = true;
                message.Attributes.Add("class", "alert alert-success");
                setValues();

            }
            else
            {
                LoadData();
                message.Text = "Please First Select MailMerge Document";
                message.Visible = true;
                message.Attributes.Add("class", "alert alert-danger");

            }
        }

        /// <summary>
        /// // generate pdf with using html text and replace with required fields in mailMerge document
        /// </summary>
        /// <param name="bodyText"></param>
        private void GeneratePdf(string bodyText)
        {
            dt = (DataTable)Session["CsvDataTable"];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                bodyText = bodyText.Replace("Contact Person Name", dt.Rows[i][0].ToString());
                FilePath = Server.MapPath("~/uploads");
                Aspose.Pdf.Generator.Pdf pdf1 = new Aspose.Pdf.Generator.Pdf();
                Aspose.Pdf.Generator.Section sec1 = pdf1.Sections.Add();
                Text title = new Text(bodyText);
                var info = new TextInfo();
                info.FontSize = 12;
                //info.IsTrueTypeFontBold = true;
                title.TextInfo = info;
                title.IsHtmlTagSupported = true;
                sec1.Paragraphs.Add(title);
                pdf1.Save(FilePath + "\\pdf\\" + dt.Rows[i][0].ToString() + ".pdf");
                bodyText = mailMergeHtml.Text;
            }
        }
        /// <summary>
        /// Sending remail function
        /// </summary>
        /// <param name="Text"></param>
        public void SendMail(string Text)
        {
            string HtmlBodyText = Text;
            dt = (DataTable)Session["CsvDataTable"];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                HtmlBodyText = HtmlBodyText.Replace("Contact Person Name", dt.Rows[i][0].ToString());
                string FilePath = Server.MapPath("~/uploads");

                // configure email with smtpserver// 

                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com", 587); //setting smtp server for gmail
                mail.From = new MailAddress("task.email.2015@gmail.com"); //set from email
                mail.To.Add(dt.Rows[i][1].ToString()); // setting required address which mail to be send
                mail.Body = "mail with attachment";
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(FilePath + "\\pdf\\" + dt.Rows[i][0].ToString() + ".pdf"); // attaching required pdf
                mail.Attachments.Add(attachment);
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("task.email.2015@gmail.com", "JawadTask");
                SmtpServer.EnableSsl = true;
                SmtpServer.Send(mail);
                HtmlBodyText = Text;
            }

            if (dt.Rows.Count > 1)
                message.Text = "" + dt.Rows.Count + " Email's Send With Required Attachment(pdf)";
            else if (dt.Rows.Count == 1)
                message.Text = "" + dt.Rows.Count + " Email' Send With Required Attachment(pdf)";

            message.Visible = true;
            message.Attributes.Add("class", "alert alert-success");
            setValues();
        }

        protected void btnTest_Click(object sender, EventArgs e)
        {
            bool connect = false;
            // checking using connected with correct smtp configration or not
            if (txtServer.Text.Trim().Equals("smtp.gmail.com"))
            {
                if (txtUserName.Text.Trim().Equals("task.email.2015@gmail.com"))
                {
                    if (txtPassword.Text.Trim().Equals("JawadTask"))
                    {
                        if (txtPort.Text.Trim().Equals("587"))
                        {
                            connect = true;
                            message.Text = "Login Successfully";
                            message.Visible = true;
                            message.Attributes.Add("class", "alert alert-success");
                            txtPassword.Text = "JawadTask";
                        }
                    }
                }
            }

            if (connect)
            {
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com", 587);
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("task.email.2015@gmail.com", "JawadTask");
                SmtpServer.EnableSsl = true;
            }
            else
            {
                setPlaceHold();
                message.Text = "Please Review ReadMe file for SMTP Settings";
                message.Visible = true;
                message.Attributes.Add("class", "alert alert-danger");
            }


        }


        // set placehold for easier understand 
        private void setPlaceHold()
        {
            if (txtServer.Text.ToString().Equals("") && txtUserName.Text.ToString().Equals("") &&
                txtPassword.Text.ToString().Equals("") && txtPort.Text.ToString().Equals(""))
            {
                txtServer.Attributes.Add("placeholder", "Server");
                txtUserName.Attributes.Add("placeholder", "User Name");
                txtPassword.Attributes.Add("placeholder", "Password");
                txtPort.Attributes.Add("placeholder", "Port");
                mailMergeHtml.Attributes.Add("placeholder", "Converted MailMerge Text In Html ");

            }

        }

        protected void btnClear_Click(object sender, EventArgs e)
        {
            txtServer.Text = string.Empty;
            txtUserName.Text = string.Empty;
            txtPassword.Text = string.Empty;
            txtPort.Text = string.Empty;
            mailMergeHtml.Text = string.Empty;
            setPlaceHold();
            message.Text = string.Empty;
            message.Visible = false;
            Session["listUsers"] = null;
            Session["listEmails"] = null;
            FilePath = string.Empty;

        }

        // send button which send email to require user email 
        protected void btnSend_Click(object sender, EventArgs e)
        {
            if (mailMergeHtml.Text != "")
                SendMail(mailMergeHtml.Text);
            else
            {
                message.Text = "Please First Upload Contact File (in CSV) & Template MailMerge";
                message.Visible = true;
                message.Attributes.Add("class", "alert alert-danger");
            }
            setPlaceHold();
        }

        private void setValues()
        {
            if (Session["listUsers"] != null)
            {
                listUsers.Items.Clear();
                foreach (ListItem Item in ((ListBox)(Session["listUsers"])).Items)
                    listUsers.Items.Add(new ListItem(Item.Text));
            }
            if (Session["listEmails"] != null)
            {
                listEmails.Items.Clear();
                foreach (ListItem Item in ((ListBox)(Session["listEmails"])).Items)
                    listEmails.Items.Add(new ListItem(Item.Text));
            }
        }

    }
}