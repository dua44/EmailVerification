using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailVerificationTester
{
    class Program
    {
        static void Main(string[] args)
        {
            EmailVerifiactionLibrary.Manager _Manager = new EmailVerifiactionLibrary.Manager();
            _Manager.StartProcess(); 
        }
    }
}
