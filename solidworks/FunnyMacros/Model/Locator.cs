using SolidWorks.Interop.sldworks;

namespace FunnyMacros.Model
{
    class Locator
    {
        public IModelDoc2 Model { set; get; }
        public IComponent2 Component { set; get; }        
    }
}
