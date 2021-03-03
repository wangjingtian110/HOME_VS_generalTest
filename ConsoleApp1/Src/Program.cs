using ConsoleApp1.Seq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("start!:");
            MainClass mainClass = new MainClass();
            mainClass.XlsTest();
            Console.WriteLine("over");
            Console.ReadLine();
        }

        public static void Main2(string[] args)
        {
            Master master1 = new Master("master1");
            Master master2 = new Master("master2");

            List<string> names = new List<string>();
            for(int i = 0; i < 3; i++)
            {
                names.Add("master" + i);
            }

            Parallel.ForEach(names, name =>
            {
                var master = new Master(name);
                for(int i = 0; i < 300; i++)
                {
                    master.send();
                }               
            });

            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine();
            Console.WriteLine("pass!!!!!!");
            Console.ReadLine();
        }
    }

    public class MainClass
    {
        public void XlsTest()
        {
            XlsMaster master = new XlsMaster();
            master.StartWork();
        }
    }
}
