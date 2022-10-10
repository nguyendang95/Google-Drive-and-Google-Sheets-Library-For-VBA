using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleApis
{
    internal class InvalidDateTimeValueException : Exception
    {
        public InvalidDateTimeValueException()
        {

        }
        public InvalidDateTimeValueException(string Message)
            : base(Message) { }
        public InvalidDateTimeValueException(string Message, Exception inner)
            : base(Message, inner) { }
    }
}
