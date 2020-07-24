using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using MailKit.Net.Smtp;
using MimeKit;
using Serilog;
using System.Data;

namespace EmailService {

    public enum AddressType {
        from,
        to,
        cc,
        bcc,
    }

    public class HTMLTableStyle {
        public string tableClass { get; set; }
    }

    public class EMailKit {

        private readonly IConfigurationRoot _config;
        //private readonly ILogger<EMailKit> _logger;

        List<string> _emailFrom;
        List<string> _emailTo;

        static Serilog.ILogger _logger = Log.ForContext(typeof(EMailKit));
        //internal static ILoggerFactory LoggerFactory { get; set; }// = new LoggerFactory();
        //internal static ILogger CreateLogger<T>() => LoggerFactory.CreateLogger<T>();
        //internal static ILogger CreateLogger(string categoryName) => LoggerFactory.CreateLogger(categoryName);

        public string smtpServer { get; set; }
        public List<string> emailFrom 
        { 
            get {
                return _emailFrom;
            }
            set {
                _emailFrom = value;
            } 
        }
        public List<string> emailTo 
        { 
            get {
                return _emailTo;
            }
            set {
                _emailTo = value;
            } 
        }
        public Dictionary<string, string> attachments = new Dictionary<string, string>();
        public string MessageHTMLBody;
        public string Subject = "";

        MimeMessage _msg;
        BodyBuilder _bodyBuilder;
        SmtpClient _client;

        //public EMailKit() { }

        public EMailKit() {
            //_logger = loggerFactory.CreateLogger<EMailKit>();
            //_logger = EmailService.CreateLogger<EMailKit>()
        }

        public EMailKit(IConfigurationRoot config, ILoggerFactory loggerFactory) {
            //_logger = loggerFactory.CreateLogger<EMailKit>();
            //_config = config;
        }

        public void AddEmailFrom(Dictionary<string, string> emailList) {
            _msg = AddEmailAddress(_msg, emailList, AddressType.from);
        }

        public void AddEmailTo(Dictionary<string, string> emailList) {
            _msg = AddEmailAddress(_msg, emailList, AddressType.to);
        }

        public void AddEmailCc(Dictionary<string, string> emailList) {
            _msg = AddEmailAddress(_msg, emailList, AddressType.cc);
        }

        public void AddEmailBcc(Dictionary<string, string> emailList) {
            _msg = AddEmailAddress(_msg, emailList, AddressType.bcc);
        }

        public MimeMessage AddEmailAddress(MimeMessage msgBody, Dictionary<string, string> emailList, AddressType addressType) {

            if (emailList == null) return msgBody;

            if (emailList.Count > 0) {
                foreach (var item in emailList) {
                    MailboxAddress addr = new MailboxAddress(item.Key, item.Value);
                    if (addressType == AddressType.from) {
                        msgBody.From.Add(addr);
                    } else if (addressType == AddressType.to) {
                        msgBody.To.Add(addr);
                    } else if (addressType == AddressType.cc) {
                        msgBody.Cc.Add(addr);
                    } else {
                        msgBody.Bcc.Add(addr);
                    }
                }
            } else {
                throw new Exception("email address [from/to] cannot be null.");
            }

            return msgBody;
        }
        
        public async Task SendEmail(Dictionary<string, string> emailFrom, Dictionary<string, string> emailTo,
                                    Dictionary<string, string> emailCc = null, Dictionary<string, string> emailBcc = null, 
                                    string templatePath="") {
            //MimeMessage msg = new MimeMessage();
            //BodyBuilder body = new BodyBuilder();
            //SmtpClient client = new SmtpClient();
            _msg = new MimeMessage();
            _bodyBuilder = new BodyBuilder();
            _client = new SmtpClient();

            AddEmailFrom(emailFrom);
            AddEmailTo(emailTo);
            AddEmailCc(emailCc);
            AddEmailBcc(emailBcc);

            if (MessageHTMLBody == null || MessageHTMLBody == "") {
                MessageHTMLBody = loadHTMLTemplate(AppContext.BaseDirectory + templatePath);
            }

            if (attachments.Count > 0) {
                foreach (var item in attachments) {
                    try {
                        /* var attachment = new MimePart(item.Value) {
                            Content = new MimeContent(File.OpenRead(item.Key), ContentEncoding.Default),
                            ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                            ContentTransferEncoding = ContentEncoding.Base64,
                            FileName = Path.GetFileName(item.Key)
                        }; */

                        _bodyBuilder.Attachments.Add(item.Key);
                        _logger.Information("Added attachment: " + item.Key);
                    } catch (Exception e) {
                        _logger.Error("Error! Cannot add attachment " + item.Key);
                        continue;
                    }
                }
            }

            _bodyBuilder.HtmlBody = MessageHTMLBody;
            _msg.Body = _bodyBuilder.ToMessageBody();
            _msg.Subject = this.Subject;

            await Task.Run(() => {
                try {

                    if (smtpServer == null || smtpServer =="") {
                        var errMsg = "Error: SMTP Server cannot be null.";
                        //_logger.LogError(errMsg);
                        throw new ArgumentNullException(errMsg);
                    }

                    _logger.Information($"Connecting to smtp server [{smtpServer}]");
                    _client.Connect(smtpServer);
                    _logger.Information("Connected. Sending mail now ..");
                    _client.Send(_msg);
                    _client.Dispose();
                } catch (Exception e) {
                    _logger.Error("Error occur on sending email. Error: " + e.Message);
                }
            });

            File.WriteAllText("output.html", MessageHTMLBody);
            _logger.Information("Email sent to: " + _msg.To.ToString());
        }

        public void replaceTemplateValue(string oldStr, string newStr) {
            MessageHTMLBody = MessageHTMLBody.Replace(oldStr, newStr);
        }

        public void loadTemplate(string path) {
            MessageHTMLBody = loadHTMLTemplate(path);
        }

        public static string loadHTMLTemplate(string fileFullPath) {
            string emailContent = "";
            StreamReader templateReader = File.OpenText(fileFullPath);
            string input = null;
            while ((input = templateReader.ReadLine()) != null) {
                emailContent = emailContent + input;
            }
            templateReader.Close();
            return emailContent;
        }

        public static string ConvertDataTableToHTML(DataTable dt, string cssClass="") {
            string html = "<table " + ((cssClass=="") ? ">" : " class=\"" + cssClass + "\">\n");

            //add header row
            html += "<thead><tr>\n";
            for (int i = 0; i < dt.Columns.Count; i++)
                html += "<th>" + dt.Columns[i].ColumnName + "</th>";
            html += "</tr></thead>\n";

            html += "<tbody>";
            //add rows
            for (int i = 0; i < dt.Rows.Count; i++) {
                html += "<tr>\n";
                for (int j = 0; j < dt.Columns.Count; j++)
                    html += "<td>" + dt.Rows[i][j].ToString() + "</td>";
                html += "</tr>\n";
            }
            html += "</tbody>";
            html += "</table>";
            return html;
        }
    }
}

