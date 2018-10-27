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

            foreach (Email email in _emailList)
            {
                if (!IsValidEmail(email.EmailAddress))
                {
                    email.VerificationStatusId = Enum.VerificationStatus.InvalidFormat;
                }
            }
        }
    }
}
