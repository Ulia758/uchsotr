using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PractikaaProizv
{
    internal class Connect
    {
        public static Database1Entities1 c;
        public static Database1Entities1 context
        {
            get
            {
                if (c == null)
                    c = new Database1Entities1();
                return c;
            }
        }
    }
}
