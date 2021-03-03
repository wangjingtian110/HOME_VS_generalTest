using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1.Seq
{
    public class Master
    {
        private static object asyncLock = new object();

        private string _name;

        private int _aux = 0;
        public int Aux { 
            get 
            {
                lock (asyncLock)
                {
                    return _aux < 255 ? _aux++ : _aux %= 255;
                }
               
            }
            private set { } }

        public Master(string name)
        {
            _name = name;
        }

        public void send()
        {
            Request request = new Request();
            request.Build(Aux);
            Console.WriteLine(string.Format("{0} has Send a message:{1}", _name, request.Msg));

        }
    }
}
