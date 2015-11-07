using SolidWorks.Interop.sldworks;

namespace FunnyMacros.Model
{
    class Base
    {
        public IFace2 Face { get; set; }
        public IComponent2 Component { get; set; }
        public double[] Box { get { return Face.GetBox() as double[]; } }
        public MathTransform Transform { get { return Component.Transform2; } }
    }
}
