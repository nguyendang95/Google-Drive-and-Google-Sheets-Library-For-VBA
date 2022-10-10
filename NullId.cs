using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleApis
{
    internal class NullIdException : Exception
    {
        public NullIdException()
        {

        }
        public NullIdException(string Message)
            : base(Message) { }
        public NullIdException(string Message, Exception inner)
            : base(Message, inner) { }
    }
}
