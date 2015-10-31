using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace FunnyMacros
{
    class Locator
    {
        public IModelDoc2 Model { set; get; }
        public IComponent2 Component { set; get; }        
    }
}
