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
            EmailNotVerified = 4,
            EmailVerified = 5
        }
    }
}
