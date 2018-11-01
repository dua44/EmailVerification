using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailVerifiactionLibrary
{
    public class Email
    {
        public int Id { get; set; }
        public string EmailAddress { get; set; }
        public EmailVerifiactionLibrary.Enum.VerificationStatus VerificationStatusId { get; set; }
        public string Verification { get; set; }
    }
}
