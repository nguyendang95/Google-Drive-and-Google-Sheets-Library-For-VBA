using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleApis
{
    internal class BlankEmailAddressOnUserGroupDomainException : Exception
    {
            public BlankEmailAddressOnUserGroupDomainException()
            {

            }
            public BlankEmailAddressOnUserGroupDomainException(string Message)
                : base(Message) { }
            public BlankEmailAddressOnUserGroupDomainException(string Message, Exception inner)
                : base(Message, inner) { }
    }
}
