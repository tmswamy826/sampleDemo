using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices.AccountManagement;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {

            WordDocReadText objword = new WordDocReadText();

            objword.FinalDistatce();
            objword.text("swamy");

            return;

            convertExceldataToJsondata objconv = new convertExceldataToJsondata();

            //string json=null;
            objconv.LoadJson();
            objconv.convertdata();
            
            Console.WriteLine("Please enter your Username");
            string uname = Console.ReadLine();
            Console.WriteLine("Please enter your Password");
            string Pwd = Console.ReadLine();
            windowsAuthentication_check(uname,Pwd);

            
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("First Line");
            builder.AppendLine();

            for (int i = 0; i <= 5; i++)
            {
                builder.AppendLine("Second Line");
            }
            builder.AppendLine("Thired line");


            string aaa = builder.ToString();
            // Display.
            //Console.Write(builder);

            // AppendLine uses \r\n sequences.
            Console.WriteLine(builder.ToString());//.Contains("\r\n"));
            Console.Read();
        }

        public static void windowsAuthentication_check( string Username,string password)
        {
            bool flag = false;


            try
            {
                using (PrincipalContext principalContext = new PrincipalContext(ContextType.Domain))
                {
                    flag = principalContext.ValidateCredentials(Username, password);
                    if(flag==true)
                    {
                        Console.WriteLine("Your given credential Matching windows authentication....");
                    }
                    else
                        Console.WriteLine("Your given credential doesnot Matching windows authentication....");
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("message :"+ ex.Message);
            }
            Console.Read();
        }
    }
}
