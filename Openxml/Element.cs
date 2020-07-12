using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Openxml
{
    public class Element
    {
        public Border border;
        public class Border
        {
            public string val;
            public string color;
            public string sz;
            public string space;
            public string shadow;
        }
        public void init()
        {
            border = new Border();
        }
    }
}
