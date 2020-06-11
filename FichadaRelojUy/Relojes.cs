using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FichadaRelojUyService
{
    public class Relojes
    {
        public int Id { get; set; }
        public string Descripcion { get; set; }
        public string DireccionIP { get; set; }
        public int Puerto { get; set; }
        public int Vaciar { get; set; }
        public int CommPassword { get; set; }
    }
}
