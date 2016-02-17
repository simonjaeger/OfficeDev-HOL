using System;

namespace Microsoft_Graph_Mail_Console_App
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            MailClient.SendMeAsync().Wait();
            Console.Read();
        }
    }
}
