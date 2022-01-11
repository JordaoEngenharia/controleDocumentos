using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Net.Mail;
using System.Reflection;

//using System.Data;
//using System.Data.SqlClient;
//using System.IO;
//using System.Drawing;
//using System.Windows.Forms;


namespace Bolinha.DAL
{
    public class Mail
    {
        public static void SendMessage(MailMessage oMessage, bool AddBCC, String NomeDe)
        {
            oMessage.From = new MailAddress("catchall@jordaoengenharia.com.br",NomeDe);
            if (AddBCC)
            {
                //oMessage.Bcc.Add(new MailAddress("marcio.americo@jordaoengenharia.com.br"));
                //oMessage.Bcc.Add(new MailAddress("doris@rj.sebrae.com.br"));
            }
            oMessage.BodyEncoding = Encoding.GetEncoding("iso-8859-1");

            SmtpClient oSmtpClient = new SmtpClient();
            oSmtpClient.Host = "mail.jordaoengenharia.com.br";
            oSmtpClient.Port = 25;
            oSmtpClient.Credentials = new System.Net.NetworkCredential("catchall@jordaoengenharia.com.br", "jordao");
            oSmtpClient.Send(oMessage);
        }
    }
}