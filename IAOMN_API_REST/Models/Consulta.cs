using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace IAOMN_API_REST.Models
{
    public class Consulta
    {
        public string[] Dimension { get; set; }
        public string[] Year { get; set; }
    }
    public class ConsultaMes: Consulta
    {
        public string Month { get; set; }
    }
}