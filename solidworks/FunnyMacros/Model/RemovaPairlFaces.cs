using SolidWorks.Interop.sldworks;

namespace FunnyMacros.Model
{
    public class RemovalPairlFaces
    {
        public IFace2 From { get; set; }
        public IFace2 To { get; set; }
        public double Distance { get; set; }
    }
}
