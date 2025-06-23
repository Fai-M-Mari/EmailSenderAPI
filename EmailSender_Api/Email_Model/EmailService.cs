using System;
using System.Net;
using System.Net.Mail;
using ClosedXML.Excel;
using System.IO;
using System.Text;

namespace EmailSender_Api.Email_Model
{
    public class EmailService
    {
        public static string SendEmailWithAttachment(string excelFilePath, string pdfFilePath, string senderEmail, string senderPassword)
        {
            // Initialize a StringBuilder for the email body
            string response = "";
            StringBuilder emailBody = new StringBuilder();
            emailBody.AppendLine("Dear Hiring Manager,");
            emailBody.AppendLine();
            emailBody.AppendLine("I hope this message finds you well. I am writing to express my strong interest in the Software Developer (.NET) position in Dubai. I will be relocating to Dubai, and I am pleased to inform you that I will be arriving in Dubai on 3th November 2023. I am eager to explore career opportunities in the region and excited about the opportunity to contribute my skills and experience to a dynamic team.");
            emailBody.AppendLine();
            emailBody.AppendLine("I hold a degree in computer science  from the University of Sindh and most recently worked as a Mid-Level Software Engineer at HBL - Habib Bank Limited. In my role, I am part of the innovation center team responsible for developing and maintaining web applications using technologies such as .NET, ASP.NET MVC, ASP.NET Core, Web Services, C#, PL/SQL, and databases.");
            emailBody.AppendLine();
            emailBody.AppendLine("With over 2 years of experience in software development, I have worked on various projects and in different roles for clients across diverse sectors. I am passionate about delivering high-quality, user-centered software and have a strong background in Agile methodologies to ensure project success. My ability to collaborate effectively in a team and lead projects sets me apart. I value collaboration, innovation, and continuous learning.");
            emailBody.AppendLine();
            emailBody.AppendLine("I have attached my resume for your reference. I would welcome the opportunity to discuss how my skills and experiences align with your client's needs. Please feel free to reach out to me at faizmuhammadmarri@gmail.com to schedule an interview.");
            emailBody.AppendLine();
            emailBody.AppendLine("Thank you for considering my application. I look forward to the possibility of joining a forward-thinking organization in Dubai and contributing to its success.");
            emailBody.AppendLine();
            emailBody.AppendLine("Sincerely,");
            emailBody.AppendLine();
            emailBody.AppendLine("Faiz Muhammad Mari\nSoftware Developer (.NET) \nPhone: (+92) 0303-2213801\nEmail: faizmuhammadmarri@gmail.com");
           
            // Load the Excel file and get the email addresses from a specific column (e.g., "Emails").
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheets.Worksheet(1); // Assuming the data is in the first worksheet
                var emailColumn = worksheet.Column("A");
                //int row = 1;
                foreach (var cell in emailColumn.CellsUsed())
                {
                    //if (row == 1)
                    //{
                    //    row = 2;
                    //}
                    //else
                    //{
                        string recipientEmail = cell.GetString();

                        using (MailMessage mail = new MailMessage())
                        {
                            mail.From = new MailAddress(senderEmail);
                            mail.To.Add(recipientEmail);
                            mail.Subject = "Applying for the Positions: Software Developer .NET";
                        string str = emailBody.ToString();
                        mail.Body = emailBody.ToString().Replace(" 2 ", " 3 ");

                            // Attach the PDF file
                            if (File.Exists(pdfFilePath))
                            {
                                Attachment attachment = new Attachment(pdfFilePath);
                                mail.Attachments.Add(attachment);
                            }

                            SmtpClient smtpClient = new SmtpClient("smtp.gmail.com")
                            {
                                Port = 587, // Replace with your SMTP server's port
                                Credentials = new NetworkCredential(senderEmail, senderPassword),
                                EnableSsl = true,
                            };

                            try
                            {
                                //smtpClient.Send(mail);
                                smtpClient.Send(mail);
                                response = $"Email sent to {recipientEmail}";
                            }
                            catch (Exception ex)
                            {
                                response = $"Failed to send email to {recipientEmail}: {ex.Message}";
                            }
                        }

                    }
                }
            //}

            return response;
        }
    }
}
