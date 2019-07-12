using System;
using AutoIt;
using System.Windows.Forms;
using System.Threading;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Program program = new Program();

                program.StartRELYApp();
                program.TestErrorReport();
                //program.MyMethod();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                MessageBox.Show("Error occured in RELY..","Error Occured" + " " + AutoItX.ErrorCode(),MessageBoxButtons.OK);
            }
        }

        public void StartRELYApp()
        {
            try
            {
                AutoItX.Run("D:\\TCS\\NBW64.exe", "D:\\", AutoItX﻿.SW_SHOW); //NBW64_7.exe
                AutoItX.Sleep(2000);

                AutoItX.Send("{Enter}");
                AutoItX.Sleep(2000);

                AutoItX.Send("KB");
                AutoItX.Sleep(2000);
				
				AutoItX.Send("{Enter}");
                AutoItX.Sleep(2000);

                /*Enter Company Code and Press ENTER Key for LOGIN to Company Menu*/
                AutoItX.Send("AUT19");
                AutoItX.Send("{Enter}");
                AutoItX.Sleep(2000);
                Console.WriteLine("LOGIN Successfully to AUT19");

                AutoItX.Send("{Enter}");
                AutoItX.Sleep(2000);

                AutoItX.Send("{Enter} {Enter}");
                AutoItX.Sleep(2000);

                AutoItX.Send("SUPER");
                AutoItX.Sleep(2000);

                AutoItX.Send("{ENTER} {Enter}");
                AutoItX.Sleep(2000);

                AutoItX.Send("{ESC}");
                AutoItX.Sleep(1000);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine();
                Console.WriteLine();

                Console.WriteLine(ex.StackTrace);
                Console.WriteLine();

                MessageBox.Show("Error occured in RELY.." + " " + AutoItX.ErrorCode());
            }           
        }

        public void TestErrorReport()
        {
            try
            {
                AutoItX.Send("R");
                AutoItX.Sleep(1000);

                AutoItX.Send("R");
                AutoItX.Sleep(2000);

                AutoItX.Send("{Enter} {Enter} {Enter}");
                AutoItX.Sleep(2000);

                AutoItX.Send("{Down} {Down} {Down} {Down} {Down}");
                AutoItX.Sleep(2000);

                AutoItX.Send("{Enter}");
                AutoItX.Sleep(2000);

                AutoItX.Send("{Right} {Right}");
                AutoItX.Sleep(2000);

                AutoItX.Send("{Enter}");
                AutoItX.Sleep(2000);

                AutoItX.Send("{Right} {Right} {Right}");
                AutoItX.Sleep(3000);

                AutoItX.Send("{Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter} {Enter}");

                //MessageBox.Show("Class files of windows ====> " + AutoItX.WinGetClassList());
                //AutoItX.Sleep(5000);

                String strGetCrWinTitle = AutoItX.WinGetTitle("[CLASS:RELY - TRIPTA Innovations Pvt. Ltd.]", "", 10);
                int flag = AutoItX.WinActivate(strGetCrWinTitle);
                Console.WriteLine("Flag is =>" + flag);

                if (flag != 0)
                {
                    MessageBox.Show("Error occured ========>" + strGetCrWinTitle);
                    Console.WriteLine("Error occured ========>" + strGetCrWinTitle);
                }
                //AutoItX.WinActive(strGetCrWinTitle);
                //MessageBox.Show("Error occured ========>" + strGetCrWinTitle);

                string errorText = AutoItX.ControlGetText(strGetCrWinTitle, "Error", "Error");

                //String errorText = AutoItX.WinGetText("[CLASS:RELY - TRIPTA Innovations Pvt. Ltd.]", "Error");

                //Console.WriteLine("Current Window title is : " + AutoItX.WinGetTitle());                
                //MessageBox.Show("Current Window title is : " + AutoItX.WinGetTitle());
                 
                MessageBox.Show("Error description ========>" + errorText);
                Console.WriteLine("Error description ========>" + errorText);

                MessageBox.Show("Description of Error page =======>" + Convert.ToString(AutoItX.WinGetHandle("[ACTIVE]")));
                Console.WriteLine("Description of Error page =======>" + Convert.ToString(AutoItX.WinGetHandle("[ACTIVE]")));

                //Int64 errorText = AutoItX.WinExists("RELY - TRIPTA Innovations Pvt. Ltd", "Error");

                //if (errorText == 0 /*"Error"*/)
                //{
                //    MessageBox.Show("Error occured in RELY.." + DateTime.Now + " " + AutoItX.ErrorCode() + AutoItX.WinGetText("[ACTIVE]"));
                //    Console.WriteLine(AutoItX.WinGetText("[ACTIVE]"));
                //    Console.ReadKey();
                //}

                /*-------------------------------------------- Error occured in Reports ---------------------------------------- */
                AutoItX.Send("{ENTER}");
                AutoItX.Sleep(2000);
                Console.Write("AFTER ERROR {ENTER} KEY IS PRESSED...!");

                AutoItX.Send("{ESC} {ESC}");
                AutoItX.Sleep(2000);
                Console.Write("AFTER ERROR {ESC} KEY IS PRESSED TWO TIMES...!");

                //MessageBox.Show("Error occured in RELY.." + " " + AutoItX.ErrorCode());

                //AutoItX.WinMinimizeAll();
                //AutoItX.WinMinimizeAllUndo();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine();
                Console.WriteLine();

                Console.WriteLine(ex.StackTrace);
                Console.WriteLine();

                MessageBox.Show("Error occured in RELY.." + ex.Message  + AutoItX.ErrorCode());
                Console.Write("Error occured in RELY.." + ex.Message + AutoItX.ErrorCode());
            }
            finally
            {
                AutoItX.WinClose();
                MessageBox.Show("RELY will be closing,now.!", "RELY will be closing,now.!", MessageBoxButtons.OK);
            }
        }

        public void MyMethod()
        {
            // Wow, this is C#!           
            AutoItX.Run﻿("notepad.exe", "", AutoItX.SW_SHOW);
            AutoItX.WinWaitActi﻿ve("Untitled");
            AutoItX.Send("I'm in notepad{Enter}{Enter}");
            AutoItX.Send("From AutoItX in C{#}{Enter}{Enter}");
            AutoItX.Send("From C{#} in AutoIt through .NET Framework{Enter}");

            //MessageBox.Show("" + AutoItX.WinGetHandleAsText());
            MessageBox.Show("All windows class are : "  + AutoItX.WinGetClassList());

            Thread.Sleep(5000);

            IntPtr winHandle = AutoItX.WinGetHandle("Untitled");
            AutoItX.WinKill(winHandle);
        }
    }
}
