using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Configuration;
using System.Diagnostics;
using QlikTelegram;

namespace QlikBot
{
    class Program
    {
        static void Main(string[] args)
        {
            Init();
            Bot qTelegramBot;
            try
            {
                qTelegramBot = new Bot();
            }
            catch (Exception e)
            {
                Console.WriteLine("Start Bot Error");
                Console.WriteLine(e.ToString());
            }

            Console.WriteLine("Type 'close' to exit.");
            bool goFlag = true;
            while (goFlag)
            {
                string c = Console.ReadLine();
                string[] parm = c.Split(' ');
                string msg;

                if (c.ToLower() == "close")
                {
                    goFlag = false;
                }
                //send message to all users
                else if (c.ToLower().StartsWith("broadcast") && parm.Length > 1)
                {
                }
                //send message to one user by id
                else if (c.ToLower().StartsWith("sendtouser") && parm.Length > 2)
                {
                }
                //send message to one user by user name
                else if (c.ToLower().StartsWith("sendtousername") && parm.Length > 2)
                {
                }
                //speak broadcast
                else if (c.ToLower().StartsWith("speak_broadcast") && parm.Length > 1)
                {
                }
                // speak to one user by id
                else if (c.ToLower().StartsWith("speak_sendtouser") && parm.Length > 2)
                {
                }
                //speak to one user by name
                else if (c.ToLower().StartsWith("speak_sendtousername") && parm.Length > 2)
                {
                }
                //change allow new user flag
                else if (c.ToLower().StartsWith("allownewusers") && parm.Length > 1)
                {
                }
            }
            end();
            //end
        }

        private static void Init()
        {
            Console.WriteLine("Chat Bot start.");
        }
        private static void end()
        {
            Console.WriteLine("Chat Bot end.");
        }
    }
}
