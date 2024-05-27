using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using AppletonEmailAPI.Models;
using System.Net.Mail;
using System.IO;
//using Spire.Xls;
using OfficeOpenXml;
using System.Web.Http.Cors;
using GemBox.Spreadsheet;
using NLog;

namespace AppletonEmailAPI.Controllers
{
    public class EmailReportController : ApiController
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();

        EmailReport[] emailData = new EmailReport[]
        {
            //new EmailReport { Id = 1, Name = "Tomato Soup", Category = "Groceries", Price = 1 },
            //new EmailReport { Id = 2, Name = "Yo-yo", Category = "Toys", Price = 3.75M },
            //new EmailReport { Id = 3, Name = "Hammer", Category = "Hardware", Price = 16.99M }
        };

        public IEnumerable<EmailReport> GetAllEmail()
        {
            return emailData;
        }

        //public IHttpActionResult GetEmail(int id)
        //{
        //    var product = emailData.FirstOrDefault((p) => p.Id == id);
        //    if (product == null)
        //    {
        //        return NotFound();
        //    }
        //    return Ok(product);
        //}
        //public void Post([FromBody]string value)
        //{
        //}

        [HttpGet]
        [Route("api/SendEmail")]
        public async System.Threading.Tasks.Task<string> SendEmailAsync()//async Task<string> sendEmail([FromBody] QuoteExportInfo quoteInfo)
        {
            
            //string binPath = string.Empty;
            string dIExportedFilesDirectory = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "Template\\";
            string templateDoc= System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "Template\\AppletonCalculatorTemplate3.xlsx";
            string outputDocumentDirectory = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "OutputDocument\\";

            // If using Professional version, put your serial key below.
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            ExcelFile ef = ExcelFile.Load(templateDoc);

            ef.Save(outputDocumentDirectory+"Convert.pdf");

            string server = "INETMAIL.EMRSN.NET";
            SmtpClient client = null;
            MailMessage mail = null;

            try
            {

                // #region send email code

                //MailAddress from = new MailAddress("flmc.salesquotes@emerson.com");
                SaveAsPdf(templateDoc);

                MailAddress from = new MailAddress("tapas.paul@emerson.com");
                MailAddress to = new MailAddress("tapas.paul@Emerson.com");
                mail = new MailMessage(from, to);

                // mail.CC.Add(new MailAddress("nakul.kadam@emerson.com"));
                //mail.CC.Add(new MailAddress("Charles.VIEIRADAROSA@Emerson.com"));

                mail.Subject = "Test Email";
                mail.Body = "This is an auto-generated email from Global Sales Quotation. Please do not reply to this email.";
                System.IO.DirectoryInfo dInfo = new System.IO.DirectoryInfo(outputDocumentDirectory);
                foreach (FileInfo file in dInfo.GetFiles())
                {
                    mail.Attachments.Add(new Attachment(outputDocumentDirectory + file.Name));
                }
                client = new SmtpClient(server);
                client.Port = 25;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.EnableSsl = false;
                //client.Credentials = new System.Net.NetworkCredential("nitin.thombre@emerson.com", "emrsn123$");
                client.Timeout = 10000;


                //WARNING - PLEASE DO NOT USE THIS BELOW LINE IN PRODUCTION. THIS IS A WORKAROUND AND NOT A COMPLETE SOLUTION. /*hck*/ - it shuts the certificate security down
                //ServicePointManager.ServerCertificateValidationCallback = delegate (object s, System.Security.Cryptography.X509Certificates.X509Certificate certificate, System.Security.Cryptography.X509Certificates.X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };


                await client.SendMailAsync(mail);


            }
            catch (Exception ex)
            {
            }

            return "success this success";
        }




        [HttpPost]
        [Route("api/SendEmailToUser")]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        public async System.Threading.Tasks.Task<string> SendEmailToUserAsync(EmailReport objEmail)//async Task<string> sendEmail([FromBody] QuoteExportInfo quoteInfo)
        {
            try
            {
                //"http://w3staging.emersonprocess.com"
                //string binPath = string.Empty;
                string dIExportedFilesDirectory = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "Template\\";
                logger.Trace(dIExportedFilesDirectory);
                string templateDoc1 = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "Template\\AppletonCalculatorTemplate.xlsx";
                string templateDoc2 = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "Template\\AppletonCalculatorTemplate2.xlsx";
                string fileName = objEmail.isProposal ? ReplaceValue(templateDoc1, objEmail) : ReplaceValue(templateDoc2, objEmail);
                //ReplaceValue(templateDoc1, objEmail);
                string templateDoc = dIExportedFilesDirectory + fileName;
                string outputDocumentDirectory = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "OutputDocument\\";
                logger.Trace(outputDocumentDirectory);

                //string server = "INETMAIL.EMRSN.NET";
                string server = "smtp.azurecomm.net";
                string smtpAuthUsername = "appleton-roi-comm|40ea5243-76d4-4b0e-a05b-ba2000e974f3|eb06985d-06ca-4a17-81da-629ab99f6505";
                string smtpAuthPassword = "ZZt8Q~06HgYiBweQkDXPFLdROy1J9s3LKFEQbcr~";
                string sender = "no-reply@emerson.com";
                string recipient = objEmail.Customer.EmailAddress;
                string bccGroup = "APPGRP.CALC@Emerson.com";
                string subject = "Appleton Light Savings Report";
                string EmailBody = "<html><body style='font-family:Arial, Helvetica, sans-serif!important'>Dear " + objEmail.Customer.CustomerName;
                EmailBody += "<br/><p>Please find attached your Appleton™ Lighting calculator savings report.</p>";
                EmailBody += "<p>The attached document details the maintenance, energy and environmental savings achieved";
                EmailBody += " by upgrading to Emerson’s Appleton™ LED luminaires.</p></br>";
                EmailBody += " <p>Learn more about Appleton LED lighting solutions by Emerson at <a href='http://www.emerson.com/en-us/automation/brands/appleton/led-lighting'>masteringled.com</a></p>";
                EmailBody += "<p>Search for a local sales representative: <a href='http://www.emersonindustrial.com/en-US/egselectricalgroup/aboutus/wheretobuy/Pages/wheretobuy.aspx'>Where To Buy</a></p>";
                EmailBody += " <p>Contact Customer Service: <a href='mailto:CustomerService.AppletonGroup@emerson.com'>CustomerService.AppletonGroup@emerson.com</a></p></br>";
                EmailBody += "<p>Best Regards</p><p>Emerson</p></br>";
                EmailBody += "<p><b>Disclaimer: </b>The information provided by the AppletonTM Lighting Retrofit Calculator is intended for use as a guide only. The calculations produced by this calculator are only estimates, and there are no guarantees that users of AppletonTM products will realize any electricity savings. The results presented by this calculator are hypothetical and may not reflect the actual performance or electricity savings at your facility.</p></body></html></br>";
                SmtpClient client = null;
                MailMessage mail = null;

                var outputDocPath=SaveAsPdf(templateDoc);
                //var outputDocPath = ConvertExcelAsMemoryStream(templateDoc);
                MailAddress from = new MailAddress(sender);
                MailAddress to = new MailAddress(recipient);
                MailAddress bcc = new MailAddress(bccGroup);
                //MailAddress bcc = new MailAddress("Tapas.paul@Emerson.com");
                mail = new MailMessage(from, to);
                if(objEmail.isBCCAllowed)
                {
                    mail.Bcc.Add(bcc);
                }             
                mail.Subject = subject;
                mail.Body = EmailBody;
                mail.IsBodyHtml = true;
                System.IO.DirectoryInfo dInfo = new System.IO.DirectoryInfo(outputDocumentDirectory);
                //foreach (FileInfo file in dInfo.GetFiles())
                //{
                    mail.Attachments.Add(new Attachment(outputDocPath));
                //}
                //mail.Attachments.Add(new Attachment(outputDocumentDirectory + fileName.Split('.')[0] + ".pdf"));
                client = new SmtpClient(server);
                //client.Port = 25;
                client.Port = 587;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                //client.EnableSsl = false;
                client.EnableSsl = true;
                client.Timeout = 10000;
                client.Credentials = new NetworkCredential(smtpAuthUsername, smtpAuthPassword);
                await client.SendMailAsync(mail);
            }
            catch (Exception ex)
            {
            }
            return "success this success";
        }

        [HttpPost]
        [Route("api/DownloadReportPDF")]
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        public byte[] DownloadReportPDF(EmailReport objEmail)//async Task<string> sendEmail([FromBody] QuoteExportInfo quoteInfo)
        {

       
            try
            {
                string dIExportedFilesDirectory = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "Template\\";
                logger.Trace(dIExportedFilesDirectory);
                string templateDoc1 = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "Template\\AppletonCalculatorTemplate.xlsx";
                string templateDoc2 = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "Template\\AppletonCalculatorTemplate2.xlsx";
                string fileName = objEmail.isProposal ? ReplaceValue(templateDoc1, objEmail) : ReplaceValue(templateDoc2, objEmail);
                string templateDoc = dIExportedFilesDirectory + fileName;
                string outputDocumentDirectory = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "OutputDocument\\";
                logger.Trace(outputDocumentDirectory);

                var outputDocPath = SaveAsPdf(templateDoc);
                return System.IO.File.ReadAllBytes(outputDocPath); 

            }
            catch (Exception ex)
            {
                return null;
            }

            
        }


        private string SaveAsPdf(string saveAsLocation)
        {
            string saveas = (saveAsLocation.Split('.')[0]) + ".pdf";
            //saveas = saveas.Replace("Template", "OutputDocument");
            saveas = saveas.Replace("AppletonCalculatorTemplate", "Lighting-Calculator-Savings-Report");
            try
            {
                //Workbook workbook = new Workbook();
                //workbook.LoadFromFile(saveAsLocation);

                ////Save the document in PDF format
                //Spire.License.LicenseProvider.SetLicenseFileName("your-license-file-name");

                //workbook.SaveToFile(saveas, Spire.Xls.FileFormat.PDF);
                // If using Professional version, put your serial key below.
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                ExcelFile ef = ExcelFile.Load(saveAsLocation);

                ef.Save(saveas);

                return saveas;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                return saveas;
            }
        }

        private string ReplaceValue(string saveAsLocation,EmailReport emailReport)
        {
            string outputFileName= "AppletonCalculatorTemplate"+new Random().Next(10000)+".xlsx";
            string exportedFileSaveAsPath = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath + "Template\\"+ outputFileName;
            var fileInformation = new System.IO.FileInfo(saveAsLocation);
            if (fileInformation.Exists)
            {
                using (ExcelPackage pck = new ExcelPackage(fileInformation))
                {
                    //ExcelWorksheet ws = pck.Workbook.Worksheets["LAYOUT1.1"];

                   OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets.First();

                    //Update customer details
                    ws.Cells["E5:G5"].Value = emailReport.Customer.CustomerName;      
                    ws.Cells["K5:M5"].Value = emailReport.Customer.EmailAddress;       
                    ws.Cells["E7:G7"].Value = emailReport.Customer.PhoneNumber;     
                    ws.Cells["K7:M7"].Value = emailReport.Customer.CompanyName;      
                    ws.Cells["E9:G9"].Value = emailReport.Customer.CompanyAddress;       
                    ws.Cells["K9:M9"].Value = emailReport.Customer.State;      
                    ws.Cells["E11:G11"].Value = emailReport.Customer.Country;      
                    ws.Cells["K11:M11"].Value = emailReport.Customer.PostalCode;

                    //Update project details
                    ws.Cells["E17:G17"].Value = emailReport.Project.ProjectName;
                    ws.Cells["K17:M17"].Value = emailReport.Project.InstallationType;
                    ws.Cells["E19:G19"].Value = emailReport.Project.Industry;
                    ws.Cells["K19:M19"].Value = emailReport.Project.Currency;
                    ws.Cells["E21:G21"].Value = emailReport.Project.EnergyCost;
                    ws.Cells["K21:M21"].Value = emailReport.Project.TimePeriod;

                    if(emailReport.isProposal)
                    {
                        //Update Appleton Data
                        //Initial Investment
                        ws.Cells["F29:F29"].Value = Convert.ToDouble(emailReport.Savings.InitialInvestment.ExistingLightingSystem);
                        ws.Cells["H29:H29"].Value = Convert.ToDouble(emailReport.Savings.InitialInvestment.ExistingLightingSystemPercentage);
                        ws.Cells["J29:J29"].Value = Convert.ToDouble(emailReport.Savings.InitialInvestment.AppletonLEDProposal);
                        ws.Cells["L29:L29"].Value = Convert.ToDouble(emailReport.Savings.InitialInvestment.AppletonLEDProposalPercentage);
                        ws.Cells["N29:N29"].Value = Convert.ToDouble(emailReport.Savings.InitialInvestment.AlternativeProposal);
                        ws.Cells["P29:P29"].Value = Convert.ToDouble(emailReport.Savings.InitialInvestment.AlternativeProposalPercentage);

                        ws.Cells["E29:E29"].Value = emailReport.Savings.InitialInvestment.ExistingLightingSystemString;
                        ws.Cells["G29:G29"].Value = emailReport.Savings.InitialInvestment.ExistingLightingSystemPercentageString;
                        ws.Cells["I29:I29"].Value = emailReport.Savings.InitialInvestment.AppletonLEDProposalString;
                        ws.Cells["K29:K29"].Value = emailReport.Savings.InitialInvestment.AppletonLEDProposalPercentageString;
                        ws.Cells["M29:M29"].Value = emailReport.Savings.InitialInvestment.AlternativeProposalString;
                        ws.Cells["O29:O29"].Value = emailReport.Savings.InitialInvestment.AlternativeProposalPercentageString;


                        ws.Cells["E31:E31"].Value = emailReport.Savings.MaintenanceCosts.ExistingLightingSystemString;
                        ws.Cells["G31:G31"].Value = emailReport.Savings.MaintenanceCosts.ExistingLightingSystemPercentageString;
                        ws.Cells["I31:I31"].Value = emailReport.Savings.MaintenanceCosts.AppletonLEDProposalString;
                        ws.Cells["K31:K31"].Value = emailReport.Savings.MaintenanceCosts.AppletonLEDProposalPercentageString;
                        ws.Cells["M31:M31"].Value = emailReport.Savings.MaintenanceCosts.AlternativeProposalString;
                        ws.Cells["O31:O31"].Value = emailReport.Savings.MaintenanceCosts.AlternativeProposalPercentageString;

                        ws.Cells["F31:F31"].Value = Convert.ToDouble(emailReport.Savings.MaintenanceCosts.ExistingLightingSystem);
                        ws.Cells["H31:H31"].Value = Convert.ToDouble(emailReport.Savings.MaintenanceCosts.ExistingLightingSystemPercentage);
                        ws.Cells["J31:J31"].Value = Convert.ToDouble(emailReport.Savings.MaintenanceCosts.AppletonLEDProposal);
                        ws.Cells["L31:L31"].Value = Convert.ToDouble(emailReport.Savings.MaintenanceCosts.AppletonLEDProposalPercentage);
                        ws.Cells["N31:N31"].Value = Convert.ToDouble(emailReport.Savings.MaintenanceCosts.AlternativeProposal);
                        ws.Cells["P31:P31"].Value = Convert.ToDouble(emailReport.Savings.MaintenanceCosts.AlternativeProposalPercentage);

                        ws.Cells["E33:E33"].Value = emailReport.Savings.EnergyCosts.ExistingLightingSystemString;
                        ws.Cells["G33:G33"].Value = emailReport.Savings.EnergyCosts.ExistingLightingSystemPercentageString;
                        ws.Cells["I33:I33"].Value = emailReport.Savings.EnergyCosts.AppletonLEDProposalString;
                        ws.Cells["K33:K33"].Value = emailReport.Savings.EnergyCosts.AppletonLEDProposalPercentageString;
                        ws.Cells["M33:M33"].Value = emailReport.Savings.EnergyCosts.AlternativeProposalString;
                        ws.Cells["O33:O33"].Value = emailReport.Savings.EnergyCosts.AlternativeProposalPercentageString;

                        ws.Cells["F33:F33"].Value = Convert.ToDouble(emailReport.Savings.EnergyCosts.ExistingLightingSystem);
                        ws.Cells["H33:H33"].Value = Convert.ToDouble(emailReport.Savings.EnergyCosts.ExistingLightingSystemPercentage);
                        ws.Cells["J33:J33"].Value = Convert.ToDouble(emailReport.Savings.EnergyCosts.AppletonLEDProposal);
                        ws.Cells["L33:L33"].Value = Convert.ToDouble(emailReport.Savings.EnergyCosts.AppletonLEDProposalPercentage);
                        ws.Cells["N33:N33"].Value = Convert.ToDouble(emailReport.Savings.EnergyCosts.AlternativeProposal);
                        ws.Cells["P33:P33"].Value = Convert.ToDouble(emailReport.Savings.EnergyCosts.AlternativeProposalPercentage);

                        ws.Cells["E35:E35"].Value = emailReport.Savings.TotalCosts.ExistingLightingSystemString;
                        ws.Cells["G35:G35"].Value = Convert.ToDouble(emailReport.Savings.TotalCosts.ExistingLightingSystem);
                        ws.Cells["I35:I35"].Value = emailReport.Savings.TotalCosts.AppletonLEDProposalString;
                        ws.Cells["K35:K35"].Value = Convert.ToDouble(emailReport.Savings.TotalCosts.AppletonLEDProposal);
                        ws.Cells["M35:M35"].Value = emailReport.Savings.TotalCosts.AlternativeProposalString;
                        ws.Cells["O35:O35"].Value = Convert.ToDouble(emailReport.Savings.TotalCosts.AlternativeProposal);

                        ws.Cells["I39:K39"].Value = emailReport.Savings.TotalSavings.AppletonTotalSaving;
                        ws.Cells["M39:O39"].Value = emailReport.Savings.TotalSavings.ProposalTotalSaving;

                        ws.Cells["I41:K41"].Value = emailReport.Savings.InitialNetInvest.AppletonTotalSaving;
                        ws.Cells["M41:O41"].Value = emailReport.Savings.InitialNetInvest.ProposalTotalSaving;

                        ws.Cells["I43:K43"].Value = emailReport.Savings.ROI.AppletonTotalSaving;
                        ws.Cells["M43:O43"].Value = emailReport.Savings.ROI.ProposalTotalSaving;

                        ws.Cells["I45:K45"].Value = emailReport.Savings.AvgSaving.AppletonTotalSaving;
                        ws.Cells["M45:O45"].Value = emailReport.Savings.AvgSaving.ProposalTotalSaving;

                        ws.Cells["I47:K47"].Value = emailReport.Savings.PaybackPeriod.AppletonTotalSaving;
                        ws.Cells["M47:O47"].Value = emailReport.Savings.PaybackPeriod.ProposalTotalSaving;
                    }
                    else
                    {
                        //Update Appleton Data
                        ws.Cells["H29:H29"].Value = Convert.ToDouble(emailReport.Savings.InitialInvestment.ExistingLightingSystem);
                        ws.Cells["J29:J29"].Value = Convert.ToDouble(emailReport.Savings.InitialInvestment.ExistingLightingSystemPercentage);
                        ws.Cells["N29:N29"].Value = Convert.ToDouble(emailReport.Savings.InitialInvestment.AppletonLEDProposal);
                        ws.Cells["P29:P29"].Value = Convert.ToDouble(emailReport.Savings.InitialInvestment.AppletonLEDProposalPercentage);

                        ws.Cells["E29:G29"].Value = emailReport.Savings.InitialInvestment.ExistingLightingSystemString;
                        ws.Cells["I29:I29"].Value = emailReport.Savings.InitialInvestment.ExistingLightingSystemPercentageString;
                        ws.Cells["K29:M29"].Value = emailReport.Savings.InitialInvestment.AppletonLEDProposalString;
                        ws.Cells["O29:O29"].Value = emailReport.Savings.InitialInvestment.AppletonLEDProposalPercentageString;


                        ws.Cells["E31:G31"].Value = emailReport.Savings.MaintenanceCosts.ExistingLightingSystemString;
                        ws.Cells["I31:I31"].Value = emailReport.Savings.MaintenanceCosts.ExistingLightingSystemPercentageString;
                        ws.Cells["K31:M31"].Value = emailReport.Savings.MaintenanceCosts.AppletonLEDProposalString;
                        ws.Cells["O31:O31"].Value = emailReport.Savings.MaintenanceCosts.AppletonLEDProposalPercentageString;

                        ws.Cells["H31:H31"].Value = Convert.ToDouble(emailReport.Savings.MaintenanceCosts.ExistingLightingSystem);
                        ws.Cells["J31:J31"].Value = Convert.ToDouble(emailReport.Savings.MaintenanceCosts.ExistingLightingSystemPercentage);
                        ws.Cells["N31:N31"].Value = Convert.ToDouble(emailReport.Savings.MaintenanceCosts.AppletonLEDProposal);
                        ws.Cells["P31:P31"].Value = Convert.ToDouble(emailReport.Savings.MaintenanceCosts.AppletonLEDProposalPercentage);


                        ws.Cells["E33:G33"].Value = emailReport.Savings.EnergyCosts.ExistingLightingSystemString;
                        ws.Cells["I33:I33"].Value = emailReport.Savings.EnergyCosts.ExistingLightingSystemPercentageString;
                        ws.Cells["K33:M33"].Value = emailReport.Savings.EnergyCosts.AppletonLEDProposalString;
                        ws.Cells["O33:O33"].Value = emailReport.Savings.EnergyCosts.AppletonLEDProposalPercentageString;

                        ws.Cells["H33:H33"].Value = Convert.ToDouble(emailReport.Savings.EnergyCosts.ExistingLightingSystem);
                        ws.Cells["J33:J33"].Value = Convert.ToDouble(emailReport.Savings.EnergyCosts.ExistingLightingSystemPercentage);
                        ws.Cells["N33:N33"].Value = Convert.ToDouble(emailReport.Savings.EnergyCosts.AppletonLEDProposal);
                        ws.Cells["P33:P33"].Value = Convert.ToDouble(emailReport.Savings.EnergyCosts.AppletonLEDProposalPercentage);


                        ws.Cells["E35:G35"].Value = emailReport.Savings.TotalCosts.ExistingLightingSystemString;
                        ws.Cells["I35:I35"].Value = Convert.ToDouble(emailReport.Savings.TotalCosts.ExistingLightingSystem);
                        ws.Cells["K35:M35"].Value = emailReport.Savings.TotalCosts.AppletonLEDProposalString;
                        ws.Cells["O35:O35"].Value = Convert.ToDouble(emailReport.Savings.TotalCosts.AppletonLEDProposal);

                        ws.Cells["I39:K39"].Value = emailReport.Savings.TotalSavings.AppletonTotalSaving;
                        //ws.Cells["M39:O39"].Value = emailReport.Savings.TotalSavings.ProposalTotalSaving;

                        ws.Cells["I41:K41"].Value = emailReport.Savings.InitialNetInvest.AppletonTotalSaving;
                        //ws.Cells["M41:O41"].Value = emailReport.Savings.InitialNetInvest.ProposalTotalSaving;

                        ws.Cells["I43:K43"].Value = emailReport.Savings.ROI.AppletonTotalSaving;
                        //ws.Cells["M43:O43"].Value = emailReport.Savings.ROI.ProposalTotalSaving;

                        ws.Cells["I45:K45"].Value = emailReport.Savings.AvgSaving.AppletonTotalSaving;
                        //ws.Cells["M45:O45"].Value = emailReport.Savings.AvgSaving.ProposalTotalSaving;

                        ws.Cells["I47:K47"].Value = emailReport.Savings.PaybackPeriod.AppletonTotalSaving;
                        //ws.Cells["M47:O47"].Value = emailReport.Savings.PaybackPeriod.ProposalTotalSaving;
                    }

                    

                    //Environmental Impact
                    ws.Cells["G68:G68"].Value = emailReport.EnvironmentalImpact.RedEnergy;
                    ws.Cells["O68:O68"].Value = emailReport.EnvironmentalImpact.SavedTree;
                    ws.Cells["G70:G70"].Value = emailReport.EnvironmentalImpact.CO2MetricTon+" Metric Ton      "+ emailReport.EnvironmentalImpact.CO2Pound + " Pound";
                    ws.Cells["O70:O70"].Value = emailReport.EnvironmentalImpact.CoalEmissionMetricTon + " Metric Ton      " + emailReport.EnvironmentalImpact.CoalEmissionPound + " Pound"; ;
                    ws.Cells["G72:G72"].Value = emailReport.EnvironmentalImpact.SavedElectricity;
                    ws.Cells["O72:O72"].Value = emailReport.EnvironmentalImpact.Car;

                    DateTime dateTime = DateTime.UtcNow.Date;
                    ws.Cells["O1:O1"].Value = dateTime.ToString("MM/dd/yyyy");


                    //Now we have to insert the quotation items to the layout template. 
                    //for first page, only 8 quotation items can be accomodated on the page

                    //starting row number/index for inserting quotation items is line/row number 15
                    int rowIndex = 15;

                    ExcelRange line = null;

  
                    


                    FileInfo fiSaveFile = new FileInfo(exportedFileSaveAsPath);
                    pck.SaveAs(fiSaveFile);

                }
            }
            return outputFileName;
        }        
    }
}
