using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleApis
{
    internal class InvalidFormatConditionIndex : Exception
    {
        public InvalidFormatConditionIndex()
        {

        }
        public InvalidFormatConditionIndex(string Message)
            : base(Message) { }
        public InvalidFormatConditionIndex(string Message, Exception inner)
            : base(Message, inner) { }
    }
}
