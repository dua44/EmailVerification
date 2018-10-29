using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace EmailVerifiactionLibrary
{
    public class Manager
    {
        List<Email> _emailList;

        public void FillData()
        {
            Email _email;
            _emailList = new List<EmailVerifiactionLibrary.Email>();
            _email = new Email();
            _email.Id = 1;
            _email.EmailAddress = "dua2004@gmail.com";
            _email.VerificationStatusId = Enum.VerificationStatus.Pending;
            _emailList.Add(_email);

            _email = new Email();
            _email.Id = 2;
            _email.EmailAddress = "dua44@hotmail.com";
            _email.VerificationStatusId = Enum.VerificationStatus.Pending;
            _emailList.Add(_email);

            _email = new Email();
            _email.Id = 3;
            _email.EmailAddress = "tcook@voiplance.com";
            _email.VerificationStatusId = Enum.VerificationStatus.Pending;
            _emailList.Add(_email);

            _email = new Email();
            _email.Id = 4;
            _email.EmailAddress = "tcooks@voiplance.com";
            _email.VerificationStatusId = Enum.VerificationStatus.Pending;
            _emailList.Add(_email);

            _email = new Email();
            _email.Id = 4;
            _email.EmailAddress = "tcooks@rizzukhantest.com";
            _email.VerificationStatusId = Enum.VerificationStatus.Pending;
            _emailList.Add(_email);

            _email = new Email();
            _email.Id = 4;
            _email.EmailAddress = "tc__ddadad@22ooks@rizzukhantest.com";
            _email.VerificationStatusId = Enum.VerificationStatus.Pending;
            _emailList.Add(_email);
        }
        //https://www.codeproject.com/Articles/12072/C-NET-DNS-query-component
        //https://social.msdn.microsoft.com/Forums/vstudio/en-US/4410b272-0564-485c-9efb-a90857c4bb12/identify-an-email-address-is-existing-or-not?forum=csharpgeneral
        public void StartProcess()
        {
            ActionResult("rizzukhan123.com");
            /*FillData();
            bool _flag = true;
            foreach (Email email in _emailList)
            {
                _flag = true;
                if (_flag)
                {
                    if (!IsValidEmail(email.EmailAddress))
                    {
                        email.VerificationStatusId = Enum.VerificationStatus.InvalidFormat;
                        _flag = false;
                    }
                }
                if (_flag)
                {
                    if (!IsValidDomain(email.EmailAddress))
                    {
                        email.VerificationStatusId = Enum.VerificationStatus.DomainNotExists;
                        _flag = false;
                    }
                }
            }

            foreach (Email email in _emailList)
            {
                Console.Write("Id : " + email.Id + " Email : " + email.EmailAddress + " Verification Status : " + email.VerificationStatusId + "\n\r");
            }
            Console.Read();
            */
        }
        public bool IsValidEmail(string _email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(_email);
                return addr.Address == _email;
            }
            catch
            {
                return false;
            }
        }
        public bool IsValidDomain(string _email)
        {
            try
            {
                System.Net.IPHostEntry ipEntry = System.Net.Dns.GetHostByName(_email.Substring(_email.IndexOf("@") + 1));
                return true;
            }
            catch (Exception ex)
            {
                if (ex.Message == "The requested name is valid, but no data of the requested type was found")
                {
                    return true;
                }
                return false;
            }
        }
        protected bool checkDNS(string host, string recType = "MX")
        {
            bool result = false;
            try
            {
                using (Process proc = new Process())
                {
                    proc.StartInfo.FileName = "nslookup";
                    proc.StartInfo.Arguments = string.Format("-type={0} {1}", recType, host);
                    proc.StartInfo.CreateNoWindow = true;
                    proc.StartInfo.ErrorDialog = false;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    proc.OutputDataReceived += (object sender, DataReceivedEventArgs e) =>
                    {
                        if ((e.Data != null) && (!result))
                            result = e.Data.StartsWith(host);
                    };
                    proc.ErrorDataReceived += (object sender, DataReceivedEventArgs e) =>
                    {
                        if (e.Data != null)
                        {
                            //read error output here, not sure what for?
                        }
                    };
                    proc.Start();
                    proc.BeginErrorReadLine();
                    proc.BeginOutputReadLine();
                    proc.WaitForExit(30000); //timeout after 30 seconds.
                }
            }
            catch
            {
                result = false;
            }
            return result;
        }
        public void ActionResult(string e)
        {
            TcpClient tClient = new TcpClient("gmail-smtp-in.l.google.com", 25);
            string CRLF = "\r\n";
            byte[] dataBuffer;
            string ResponseString;

            NetworkStream netStream = tClient.GetStream();
            StreamReader reader = new StreamReader(netStream);
            ResponseString = reader.ReadLine();
            /* Perform HELO to SMTP Server and get Response */
            dataBuffer = BytesFromString("HELO " + CRLF);
            netStream.Write(dataBuffer, 0, dataBuffer.Length);
            ResponseString = reader.ReadLine();
            dataBuffer = BytesFromString("MAIL FROM:<mycrmcheck@gmail.com>" + CRLF);
            netStream.Write(dataBuffer, 0, dataBuffer.Length);
            ResponseString = reader.ReadLine();
            /* Read Response of the RCPT TO Message to know from google if it exist or not */

            dataBuffer = BytesFromString("RCPT TO:<" + e.ToString().Trim() + ">" + CRLF);
            netStream.Write(dataBuffer, 0, dataBuffer.Length);
            ResponseString = reader.ReadLine();
            if (GetResponseCode(ResponseString) == 550)
            {
                Console.Write("Mai Address Does not Exist !\n\r");
                Console.Write("Original Error from Smtp Server :" + ResponseString);
            }
            /* QUITE CONNECTION */
            Console.Write("Email Id Existing !");
            dataBuffer = BytesFromString("QUITE" + CRLF);
            netStream.Write(dataBuffer, 0, dataBuffer.Length);
            tClient.Close();            
        }

        public byte[] BytesFromString(string str)
        {
            return Encoding.ASCII.GetBytes(str);
        }
        public int GetResponseCode(string ResponseString)
        {
            return int.Parse(ResponseString.Substring(0, 3));
        }
    }
}