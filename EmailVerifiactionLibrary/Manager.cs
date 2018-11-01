using LumiSoft.Net.DNS.Client;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
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
        public List<Email> ReadExcelData(string FilePath)
        {
            Email _email;
            List<Email> _myemaillist = new List<Email>();
            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 2; row <= rowCount; row++)
                {
                    _email = new Email();
                    _email.EmailAddress = worksheet.Cells[row, 1].Value.ToString().Trim();
                    _email.Id = row;
                    _email.VerificationStatusId = Enum.VerificationStatus.Pending;
                    _myemaillist.Add(_email);
                }
            }
            return _myemaillist;
        }
        //https://www.codeproject.com/Articles/12072/C-NET-DNS-query-component
        //https://social.msdn.microsoft.com/Forums/vstudio/en-US/4410b272-0564-485c-9efb-a90857c4bb12/identify-an-email-address-is-existing-or-not?forum=csharpgeneral
        public void StartProcess()
        {
            List<Email> _finalemaillist = new List<Email>();
            List<Email> _emailList = new List<Email>();
            bool _flag = false;
            _emailList = ReadExcelData(@"C:\Users\rizwanahmad\Downloads\1stLot.xlsx");
            var _domain = _emailList.Select(x => x.EmailAddress.Split('@').GetValue(1)).Distinct();

            foreach (string s in _domain)
            {
                _flag = false;
                _flag = IsValidDomain(s);
                if (!_flag)
                {
                    var _invaliddomainemail = _emailList.Where(x => x.EmailAddress.Contains(s)).ToList();
                    foreach (Email e in _invaliddomainemail)
                    {
                        e.VerificationStatusId = Enum.VerificationStatus.DomainNotExists;
                        _finalemaillist.Add(e);
                    }
                }
                else
                {
                    _flag = IsValidMXRecord(s);
                    if (!_flag)
                    {
                        var _invalidmxrecordemail = _emailList.Where(x => x.EmailAddress.Contains(s)).ToList();
                        foreach (Email e in _invalidmxrecordemail)
                        {
                            e.VerificationStatusId = Enum.VerificationStatus.MXRecordNotFound;
                            _finalemaillist.Add(e);
                        }
                    }
                    else
                    {
                        var _formatemail = _emailList.Where(x => x.EmailAddress.Contains(s)).ToList();
                        foreach (Email e in _formatemail)
                        {
                            if (IsValidEmail(e.EmailAddress))
                            {
                                e.VerificationStatusId = Enum.VerificationStatus.EmailVerified;
                            }
                            else
                            {
                                e.VerificationStatusId = Enum.VerificationStatus.InvalidFormat;
                            }
                            _finalemaillist.Add(e);
                        }
                    }
                }
            }

            foreach (Email email in _finalemaillist)
            {
                Console.Write("Id : " + email.Id + " Email : " + email.EmailAddress + " Verification Status : " + email.VerificationStatusId + "\n\r");
            }
            Console.Read();
            
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
        public bool IsValidDomain(string _domain)
        {
            try
            {
                System.Net.IPHostEntry ipEntry = System.Net.Dns.GetHostByName(_domain);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        public bool IsValidMXRecord(string _domain)
        {
            try
            {
                Dns_Client.DnsServers = new string[] { "8.8.8.8" };
                Dns_Client.UseDnsCache = true;
                using (Dns_Client dns = new Dns_Client())
                {
                    DnsServerResponse reponse = null;
                    reponse = dns.Query(_domain, LumiSoft.Net.DNS.DNS_QType.MX);
                    if (((LumiSoft.Net.DNS.DNS_rr_MX)reponse.Answers[0]).Host != null)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                return false;
            }
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
        public static string SaveToExcel(DataTable dt, string FileName, int AgentCode, int DataCategoryId, int DataGenerationRequestId)
        {


            /*Set up work book, work sheets, and excel application*/
            string pathfordb = string.Empty;
            /*Set up work book, work sheets, and excel application*/
            //Microsoft.Office.Interop.Excel.Application oexcel = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                if (dt == null || dt.Rows.Count == 0)
                {
                    dt.Columns.Add("Message");
                    dt.Rows.Add("No Records Found!");
                }
                DataColumnCollection columns = dt.Columns;
                if (DataCategoryId == Convert.ToInt32(Constants.DataCategory.UpSell))
                {
                    if (columns.Contains("Email_Address"))
                        dt.Columns.Remove("Email_Address");
                    if (columns.Contains("Email"))
                        dt.Columns.Remove("Email");
                    if (columns.Contains("Phone_Number"))
                        dt.Columns.Remove("Phone_Number");
                }
                if (DataCategoryId == Convert.ToInt32(Constants.DataCategory.Email))
                {
                    if (columns.Contains("Phone_Number"))
                        dt.Columns.Remove("Phone_Number");
                }
                if (DataCategoryId == Convert.ToInt32(Constants.DataCategory.Sms))
                {
                    if (columns.Contains("Email_Address"))
                        dt.Columns.Remove("Email_Address");
                    if (columns.Contains("Email"))
                        dt.Columns.Remove("Email");
                }

                // object misValue = System.Reflection.Missing.Value;
                //Microsoft.Office.Interop.Excel.Workbook obook = oexcel.Workbooks.Add();
                //Microsoft.Office.Interop.Excel.Worksheet osheet = new Microsoft.Office.Interop.Excel.Worksheet();
                //  obook.Worksheets.Add(misValue);
                //osheet = (Microsoft.Office.Interop.Excel.Worksheet)obook.Sheets["Sheet1"];
                //int colIndex = 0;
                //int rowIndex = 1;
                //foreach (DataColumn dc in dt.Columns)
                //{
                //    colIndex++;
                //    osheet.Cells[1, colIndex] = dc.ColumnName;
                //}
                //foreach (DataRow dr in dt.Rows)
                //{
                //    rowIndex++;
                //    colIndex = 0;

                //    foreach (DataColumn dc in dt.Columns)
                //    {
                //        colIndex++;
                //        osheet.Cells[rowIndex, colIndex] = dr[dc.ColumnName];
                //    }
                //}
                //osheet.Columns.AutoFit();
                clsGeneral objclsgeneral = new DataGenerationLibrary.clsGeneral();
                DataSet dsData = new DataSet();
                string targetFolder = ConfigurationManager.AppSettings["DataFolderName"].ToString();
                string path = AppDomain.CurrentDomain.BaseDirectory + targetFolder + "\\" + AgentCode + "\\";
                pathfordb = targetFolder + "\\" + AgentCode + "\\" + FileName + ".xlsx";
                if (!Directory.Exists(path)) { Directory.CreateDirectory(path); }
                string filepath = path + FileName + ".xlsx";
                DataTable _NewDataTable = dt.Copy();
                dsData.Tables.Add(_NewDataTable);
                objclsgeneral.GenerateExcel2007(filepath, dsData);
                //Release and terminate excel
                //obook.SaveAs(filepath);
                //obook.Close();
                //oexcel.Quit();
                //GC.Collect();
            }
            catch (Exception ex)
            {
                //oexcel.Quit();
                Logging.WriteToLog(Logging.GenerateDefaultLogFileName("DataGenerationService"), "Error While Putting Data in Excel >> MethodName=SaveToExcel >> DataGenerationRequestId=" + DataGenerationRequestId + ">> Exception =" + ex.Message);
                Methods.UpdateDataRequestStatus(Convert.ToInt32(Constants.ServiceStatus.Error), DataGenerationRequestId);
                pathfordb = string.Empty;
            }
            return pathfordb;
        }
    }
}