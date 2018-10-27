using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailVerifiactionLibrary
{
    public class Manager
    {
        List<Email> _emailList;
        bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
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
        }
        public void main()
        {
            FillData();
            bool _flag = true;
            foreach (Email email in _emailList)
            {
                if (_flag)
                {
                    if (!IsValidEmail(email.EmailAddress))
                    {
                        email.VerificationStatusId = Enum.VerificationStatus.InvalidFormat;
                        _flag = false;
                    }
                    if (_flag)
                    {
                        //https://www.codeproject.com/Articles/12072/C-NET-DNS-query-component
                        //https://social.msdn.microsoft.com/Forums/vstudio/en-US/4410b272-0564-485c-9efb-a90857c4bb12/identify-an-email-address-is-existing-or-not?forum=csharpgeneral
                    }
                }
            }
        }
    }
}
