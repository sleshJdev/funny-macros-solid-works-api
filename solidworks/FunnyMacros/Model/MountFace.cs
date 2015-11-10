using SolidWorks.Interop.sldworks;
using FunnyMacros.Util;

namespace FunnyMacros.Model
{
    class MountFace : Model
    {
        public MountFace(IMathUtility mathUtility) : base(mathUtility)
        {
        }

        public IFace2 Face { get; set; }

        public Box FaceLBox
        {
            get { return new Box(Face.GetBox()); }
        }

        public Box FaceGBox
        {
            get { return new Box(Helper.ApplyTransform(Transform, FaceLBox.CoordinatesCorners)); }
        }

        public double Radius
        {
            get { return Face.IGetSurface().CylinderParams[6]; }
        }
    }
}
