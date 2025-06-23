using EmailSender_Api.Email_Model;
using Microsoft.AspNetCore.Mvc;
using System;
 

namespace EmailSender_Api.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class Send_EmailController : ControllerBase
    {
        [HttpGet("sendemail")]
        public IActionResult SendEmail()
        {
            try
            {
                // Code for sending an email here
                string excelFilePath = "F:\\Sindhi Developer\\DubaiHR_Emails.xlsx";
                string excelFpdfFilePathilePath = "F:\\Sindhi Developer\\Faiz Muhammad Mari CV.pdf";
                string senderEmail = "faizmuhammadmarri@gmail.com";
                string senderPassword = "ntwt lnkv pqan sdbo";

               string Response =  EmailService.SendEmailWithAttachment(excelFilePath, excelFpdfFilePathilePath, senderEmail, senderPassword);
                // If the email is sent successfully, return a success response
                return Ok(Response);
            }
            catch (Exception ex)
            {
                // If there is an error, return an error response
                return BadRequest($"Failed to send email: {ex.Message}");
            }
        }
    }
}
