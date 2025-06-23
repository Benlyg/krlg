using Report.Repository;
using Report.Service;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Report
{
    static class Program
    {
        private static LogWriter _logWriter = new LogWriter();
        readonly static private IPrintLabel _print = new PrintLabel_REPO();
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            _logWriter.LogWrite(string.Format("app start {0}", DateTime.Now));
            //IPrintLabel _pritn = new Print_Label();
            if (args == null)
            {
                _logWriter.LogWrite(string.Format("args == null"));
            }
            else if (args.Length == 3)
            {
                _logWriter.LogWrite(string.Format("args.Length == 3"));
                _logWriter.LogWrite(string.Format("pathjson = {0} , pathRDLC = {1} , printName = {2}", args[0], args[1], args[2]));
                string pathjson = args[0];
                string pathRDLC = args[1];
                string printername = args[2];
                _print.Print_label(pathjson, pathRDLC, printername);
            }
            else if (args.Length == 4)
            {
                if (args[3] == "PDF")
                {
                    _logWriter.LogWrite(string.Format("args.Length == 4"));
                    _logWriter.LogWrite(string.Format("pathjson = {0} , pathRDLC = {1} , printName = {2} , type = {3}", args[0], args[1], args[2], args[3]));
                    string pathjson = args[0];
                    string pathRDLC = args[1];
                    string printername = args[2];
                    string type = args[3];
                    _print.Print_label_PDF(pathjson, pathRDLC, printername);
                }
            }
            else
            {
                _logWriter.LogWrite(string.Format("args.Length = {0}", args.Length));
            }

            _logWriter.LogWrite(string.Format("app stop {0}", DateTime.Now));
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
        }
    }
}
