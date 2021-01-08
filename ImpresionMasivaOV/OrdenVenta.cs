using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImpresionMasivaOV
{

    public class Model
    {
        public List<OV> listOV { get; set; }
    }
    public class OV
    {
        public  oEncabezado Encabezado { get; set; }
        public List<oDetalle> Detalle { get; set; }
    }

    public class oEncabezado
    {
        public int docentry { get; set; }
        public int docnum { get; set; }
        public string ruta { get; set; }
        public string cardCode { get; set; }
        public string cardName { get; set; }
        public string dirDespacho { get; set; }
        public string comuna { get; set; }
        public string ciudad { get; set; }
        public double docTotal { get; set; }
        public string CorrelativoERP { get; set; }
        public DateTime fechaDespacho { get; set; }
    }

    public class oDetalle
    {
        public string codigo { get; set; }
        public string descripcion { get; set; }
        public double cantiad { get; set; }
        public double precio { get; set; }
        public double total { get; set; }
        public double docentry { get; set; }
        public double lineNum { get; set; }
    }
}
