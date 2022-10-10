using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleApis
{
    internal class InvalidSheetIndex : Exception
    {
        public InvalidSheetIndex()
        {

        }
        public InvalidSheetIndex(string Message)
            : base(Message) { }
        public InvalidSheetIndex(string Message, Exception inner)
            : base(Message, inner) { }
    }
}
