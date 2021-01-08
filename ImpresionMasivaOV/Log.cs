using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImpresionMasivaOV
{
   public class Log
    {

        public  void AddLog(String Mensaje)
        {
            StreamWriter Arch;
            //Exe: String = 
            String sPath = Path.GetDirectoryName(this.GetType().Assembly.Location);
            String NomArch;
            String NomArchB;
            NomArch = "\\VDLog_" + String.Format("{0:yyyy-MM-dd}", DateTime.Now) + ".log";
            Arch = new StreamWriter(sPath + NomArch, true);
            NomArchB = sPath + "\\VDLog_" + String.Format("{0:yyyy-MM-dd}", DateTime.Now.AddDays(-1)) + ".log";
            //Elimina archivo del dia anterior
            //if (System.IO.File.Exists(NomArchB))
            //    System.IO.File.Delete(NomArchB);

            try
            {
                Arch.WriteLine(String.Format("{0:dd-MM-yyyy HH:mm:ss}", DateTime.Now) + " " + Mensaje);
            }
            finally
            {
                Arch.Flush();
                Arch.Close();
            }

        }

    }
}
