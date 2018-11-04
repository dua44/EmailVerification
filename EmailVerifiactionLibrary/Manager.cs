using LumiSoft.Net.DNS.Client;
using OfficeOpenXml;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EmailVerifiactionLibrary
{
    public class Manager
    {
        public string _FileDirectory = ConfigurationManager.AppSettings["FileDirectory"].ToString();
        public List<Email> ReadExcelData(string FilePath)
        {
            Email _email;
            List<Email> _myemaillist = new List<Email>();
            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheets worksheet = package.Workbook.Worksheets;
                foreach (ExcelWorksheet sheet in worksheet)
                {
                    int colCount = sheet.Dimension.End.Column;  //get Column Count
                    int rowCount = sheet.Dimension.End.Row;     //get row count
                    for (int row = 2; row <= rowCount; row++)
                    {
                        if (sheet.Cells[row, 1].Value != null)
                        {
                            if (sheet.Cells[row, 1].Value.ToString().Trim().IndexOf("@") > 0)
                            {
                                _email = new Email();
                                _email.EmailAddress = sheet.Cells[row, 1].Value.ToString().Trim();
                                _email.Id = _myemaillist.Count + 1;
                                _email.VerificationStatusId = Enum.VerificationStatus.Pending;
                                _email.Verification = Enum.VerificationStatus.Pending.ToString();
                                _myemaillist.Add(_email);
                            }
                        }
                    }
                }

            }
            return _myemaillist;
        }
        public void StartProcess(string _inputFileName)
        {
            ConcurrentBag<Email> _finalemaillist = new ConcurrentBag<Email>();
            List<Email> _emailList = new List<Email>();
            bool _flag = false;
            _emailList = ReadExcelData(_FileDirectory + _inputFileName);
            var _domain = _emailList.Select(x => x.EmailAddress.Split('@').GetValue(1)).Distinct();
            var dict = _emailList.GroupBy(x => x.EmailAddress.Split('@').GetValue(1).ToString())
                        .Select(x => new
                        {
                            Domain = x.Key,
                            Data = x.ToList()
                        })


               .ToDictionary(x => x.Domain, x => x.Data);
            //foreach (var s in dict)
             Parallel.ForEach(dict, new ParallelOptions { MaxDegreeOfParallelism = 2 } , s =>
            {
                _flag = false;
                _flag = IsValidDomain(s.Key);
                if (!_flag)
                {
                    s.Value.AsParallel().ForAll(c =>
                    {

                        //c.VerificationStatusId = Enum.VerificationStatus.DomainNotExists;
                        //c.Verification = Enum.VerificationStatus.DomainNotExists.ToString();
                        //_finalemaillist.Add(c);

                        _finalemaillist.Add(new Email()
                        {
                            Id = c.Id,
                            EmailAddress = c.EmailAddress,
                            VerificationStatusId = Enum.VerificationStatus.DomainNotExists,
                            Verification = Enum.VerificationStatus.DomainNotExists.ToString()

                        });
                    });

                    //foreach (Email e in _invaliddomainemail)
                    //{
                    //    e.VerificationStatusId = Enum.VerificationStatus.DomainNotExists;
                    //    e.Verification = Enum.VerificationStatus.DomainNotExists.ToString();
                    //    if (!_finalemaillist.Any(x => x.EmailAddress == e.EmailAddress))
                    //    {
                    //        _finalemaillist.Add(e);
                    //    }
                    //}
                    // _finalemaillist.AddRange(_invaliddomainemail);

                    // s.Value.AsParallel().ForAll(x => { });
                }
                else
                {
                    Console.WriteLine($"{Enum.VerificationStatus.MXRecordNotFound.ToString()} {s.Key} {s.Value.Count}");
                    _flag = IsValidMXRecord(s.Key);
                    if (!_flag)
                    {
                        s.Value.AsParallel().ForAll(c =>
                        {
                            //c.VerificationStatusId = Enum.VerificationStatus.MXRecordNotFound;
                            //c.Verification = Enum.VerificationStatus.MXRecordNotFound.ToString();

                            _finalemaillist.Add(new Email()
                            {
                                Id = c.Id,
                                EmailAddress = c.EmailAddress,
                                VerificationStatusId = Enum.VerificationStatus.MXRecordNotFound,
                                Verification = Enum.VerificationStatus.MXRecordNotFound.ToString()

                            });

                            // _finalemaillist.Add(c);
                        });
                        //foreach (Email e in _invalidmxrecordemail)
                        //{
                        //    e.VerificationStatusId = Enum.VerificationStatus.MXRecordNotFound;
                        //    e.Verification = Enum.VerificationStatus.MXRecordNotFound.ToString();
                        //    if (!_finalemaillist.Any(x => x.EmailAddress == e.EmailAddress))
                        //    {
                        //        _finalemaillist.Add(e);
                        //    }
                        //}
                        // s.Value.AsParallel().ForAll(x => { });
                        // _finalemaillist.AddRange(_invalidmxrecordemail);
                    }
                    else
                    {
                        //var _formatemail = new ConcurrentBag<Email>();

                        //  s.Value.AsParallel().ForAll(e =>
                        Console.WriteLine( $"{Enum.VerificationStatus.EmailVerified.ToString()} {s.Key} {s.Value.Count}");
                        Parallel.ForEach(s.Value, e =>
                        {
                              if (IsValidEmail(e.EmailAddress))
                              {

                                  _finalemaillist.Add(new Email()
                                  {
                                      Id = e.Id,
                                      EmailAddress = e.EmailAddress,
                                      VerificationStatusId = Enum.VerificationStatus.EmailVerified,
                                      Verification = Enum.VerificationStatus.EmailVerified.ToString()

                                  });
                              }
                              else
                              {

                                  _finalemaillist.Add(new Email()
                                  {
                                      Id = e.Id,
                                      EmailAddress = e.EmailAddress,
                                      VerificationStatusId = Enum.VerificationStatus.InvalidFormat,
                                      Verification = Enum.VerificationStatus.InvalidFormat.ToString()

                                  });

                              }



                          });
                        //s.Value.AsParallel().ForAll(x => {});

                    }
                }
                Console.WriteLine("Total Object :" + _finalemaillist.Count);
            }
            );
            SaveToExcel(_finalemaillist.ToList(), _FileDirectory, "Generated-" + _inputFileName);
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
        public static string SaveToExcel(List<Email> _emailList, string _directory, string _fileName)
        {
            DataTable dt = ToDataTable<Email>(_emailList);
            /*Set up work book, work sheets, and excel application*/
            string pathfordb = string.Empty;
            /*Set up work book, work sheets, and excel application*/
            //Microsoft.Office.Interop.Excel.Application oexcel = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                DataColumnCollection columns = dt.Columns;
                clsGeneral objclsgeneral = new clsGeneral();
                DataSet dsData = new DataSet();
                string filepath = _directory + _fileName;
                DataTable _NewDataTable = dt.Copy();
                dsData.Tables.Add(_NewDataTable);
                objclsgeneral.GenerateExcel2007(filepath, dsData);
            }
            catch (Exception ex)
            {
                pathfordb = string.Empty;
            }
            return pathfordb;
        }
        public static DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Defining type of data column gives proper data table 
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name, type);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }
    }
}