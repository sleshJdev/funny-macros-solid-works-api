using FunnyMacros.Util;
using SolidWorks.Interop.sldworks;

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

        public Vector CenterFaceGBox
        {
            get { return Helper.CenterBox(FaceGBox); }
        }
    }
}
