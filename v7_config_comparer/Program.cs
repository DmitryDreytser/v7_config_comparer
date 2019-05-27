using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using v7MetaData;

namespace v7_config_comparer
{
    static class Program
    {
        public static OleStorage.TaskItem First;
        public static OleStorage.TaskItem Second;

        public static string FirstFileName = string.Empty;
        public static string SecondFileName = string.Empty;

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [STAThread]
        static void Main(string[] arg)
        {
            if (arg.Length != 2)
            {
                var handle = GetConsoleWindow();
                ShowWindow(handle, 0);

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            else
            {
                FirstFileName = arg[0];
                SecondFileName = arg[1];

                if (File.Exists(FirstFileName) && File.Exists(SecondFileName))
                {
                    First = new OleStorage.TaskItem(FirstFileName);
                    Second = new OleStorage.TaskItem(SecondFileName);

                    Console.WriteLine(First.CompareWith(Second, true, true).Replace("\n", "\r\n"));
                }
                else
                {
                    Console.WriteLine("Одна из конфигураций не существует!");
                }
                //Console.ReadKey();
            }


        }


    }
}
