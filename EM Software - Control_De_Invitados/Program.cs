using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.SqlServerCe;

using System.Threading;
using System.IO.Log;
using System.IO;
using System.Text;

namespace EM_Software___Control_De_Invitados
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        //public DataTable dt;
        //public SqlCeDataAdapter da;
        //public SqlCeConnection PathBD;
        //public DataRow row;
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);    
            Application.Run(new Form1());
        }    

        static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            using (FileRecordSequence record = new FileRecordSequence("application.log", FileAccess.Write))
            {

                string message = string.Format("[{0}]Message::{1} StackTrace:: {2}", DateTime.Now,
                                                                                    e.Exception.Message,
                                                                                    e.Exception.StackTrace);

                record.Append(CreateData(message), SequenceNumber.Invalid,
                                                    SequenceNumber.Invalid,
                                                    RecordAppendOptions.ForceFlush);
            }
       }

        private static IList<ArraySegment<byte>> CreateData(string str)
        {
            Encoding enc = Encoding.Unicode;

            byte[] array = enc.GetBytes(str);

            ArraySegment<byte>[] segments = new ArraySegment<byte>[1];
            segments[0] = new ArraySegment<byte>(array);

            return Array.AsReadOnly<ArraySegment<byte>>(segments);
        }
    }




}
