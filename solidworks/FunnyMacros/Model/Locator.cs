using FunnyMacros.Util;
using SolidWorks.Interop.sldworks;
using System;

namespace FunnyMacros.Model
{
    class Locator
    {
        public IModelDoc2 Model { set; get; }
        public IComponent2 Component { set; get; }
        public double[] Box { get { return Component.GetBox(true, true) as double[]; } }
        public MathTransform Transform { get { return Component.Transform2; } }
        public IEquationMgr EquationManager { get { return Model.GetEquationMgr(); } }
        public IBody2 Body { get { return Component.GetBody(); } }
        public IFace2[] Faces
        {
            get
            {
                return Array.ConvertAll(Body.GetFaces(), new Converter<object, IFace2>((o) => { return o as IFace2; }));
            }
        }
        public IFace2 GetBottom(IMathUtility mathUtility)
        {
            return FindVecticalExtremeFaceWithNormal(mathUtility, -1);
        }

        public IFace2 GetTop(IMathUtility mathUtility)
        {
            return FindVecticalExtremeFaceWithNormal(mathUtility, 1);
        }

        private IFace2 FindVecticalExtremeFaceWithNormal(IMathUtility mathUtility, int direction)
        {
            IFace2 extremeFace = null;
            double extremeY = direction * double.MinValue;
            IMathVector targetNormalVector = mathUtility.CreateVector(new double[] { 0.0, direction, 0.0});
            for (int i = 0; i < Faces.Length; ++i)
            {
                IFace2 face = Faces[i];
                if (face.IGetSurface().IsPlane())
                {
                    double[] normal = MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, Transform, face.Normal);
                    IMathVector normalVector = mathUtility.CreateVector(normal);
                    if (MiscUtil.IsCoDirectional(normalVector, targetNormalVector) || 
                        MiscUtil.IsCoDirectional(MiscUtil.Negative(mathUtility, normalVector), targetNormalVector))
                    {
                        double[] center = MiscUtil.GetCenterOf(face.GetBox() as double[]);
                        if (direction == 1 ? center[1] > extremeY : center[1] < extremeY)
                        {
                            extremeY = center[1];
                            extremeFace = face;
                        }
                    }
                }
            }
            return extremeFace;
        }
    }
}
