using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1.Seq
{
    public class Request
    {

        public string Msg { get; set; }

        public void Build(int num)
        {
            Msg =  "my seq is " + num;
        }
    }
}
