using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailVerifiactionLibrary
{
    public class Enum
    {
        public enum VerificationStatus
        {
            Pending = 1,
            InvalidFormat = 2,
            DomainNotExists = 3,
            MXRecordNotFound = 4,
            EmailNotVerified = 5,
            EmailVerified = 6
        }

        private enum SMTPResponse : int
        {
            CONNECT_SUCCESS = 220,
            GENERIC_SUCCESS = 250,
            DATA_SUCCESS = 354,
            QUIT_SUCCESS = 221
        }
    }
}
