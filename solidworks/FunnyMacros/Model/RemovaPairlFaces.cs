using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FunnyMacros.Model
{
    class RemovalPairlFaces
    {
        public IFace2 From { get; set; }
        public IFace2 To { get; set; }
        public double Distance { get; set; }
    }
}
