using FunnyMacros.Util;
using SolidWorks.Interop.sldworks;
using System;

namespace FunnyMacros.Model
{
    class Locator : Model
    {
        public Locator(IMathUtility mathUtility) : base(mathUtility)
        {
        }

        public IModelDoc2 Model { set; get; }
        public IEquationMgr EquationManager { get { return Model.GetEquationMgr(); } }
        
        public IFace2 BottomFace { get { return FindVecticalExtremeFaceWithNormal(-1); } }
        public IFace2 TopFace { get { return FindVecticalExtremeFaceWithNormal(1); } }

        private IFace2 FindVecticalExtremeFaceWithNormal(int direction)
        {
            IFace2 extremeFace = null;
            Vector extreme = new Vector(0.0, direction * double.MinValue, 0.0);
            IMathVector targetNormalVector = MathUtility.CreateVector(new double[] { 0.0, direction, 0.0 });
            for (int i = 0; i < Faces.Length; ++i)
            {
                IFace2 face = Faces[i];
                if (face.IGetSurface().IsPlane())
                {
                    double[] normal = Helper.Instance.ApplyTransform(Transform, face.Normal);
                    IMathVector normalVector = MathUtility.CreateVector(normal);
                    if (Helper.IsCoDirectional(normalVector, targetNormalVector) ||
                        Helper.IsCoDirectional(Helper.Negative(normalVector), targetNormalVector))
                    {
                        Vector center = Helper.CenterBox(new Box(face.GetBox()));
                        if (direction == 1 ? center.Y > extreme.Y : center.Y < extreme.Y)
                        {
                            extreme.Y = center.Y;
                            extremeFace = face;
                        }
                    }
                }
            }

            return extremeFace;
        }
    }
}
