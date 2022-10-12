using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleApis
{
    internal class SheetNotFound : Exception
    {
        public SheetNotFound()
        {

        }
        public SheetNotFound(string Message)
            : base(Message) { }
        public SheetNotFound(string Message, Exception inner)
            : base(Message, inner) { }
    }
}
