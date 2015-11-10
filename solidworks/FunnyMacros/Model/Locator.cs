using SolidWorks.Interop.sldworks;
using FunnyMacros.Util;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System;

namespace FunnyMacros.Model
{
    class Locator : Model
    {
        public Locator(IMathUtility mathUtility) : base(mathUtility)
        {
        }

        public string FullPath { get; set;}

        public IModelDoc2 Model
        {
            get { return Component.GetModelDoc2(); }
        }
        public IEquationMgr EquationManager
        {
            get { return Model.GetEquationMgr(); }
        }
        
        public IFace2 BottomFace
        {
            get { return FindVecticalExtremeFaceWithNormal(-1); }
        }

        public IFace2 TopFace
        {
            get { return FindVecticalExtremeFaceWithNormal(1); }
        }

        public void Scale(short type, bool uniform, double factorX, double factorY, double factorZ)
        {
            Model.FeatureManager.InsertScale(type, uniform, factorX, factorY, factorZ);
        }

        private bool IsHorizontal(IMathVector normal)
        {
            double[] normalVecor = normal.ArrayData;
            return Math.Abs(normalVecor[1]) > Math.Abs(normalVecor[0]) &&
                   Math.Abs(normalVecor[1]) > Math.Abs(normalVecor[2]) &&
                   Math.Abs(normalVecor[1]) > 0.5;
        }

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
                    if (IsHorizontal(normalVector) || IsHorizontal(Helper.Negative(normalVector)))
                    {
                        Vector center = new Box(Helper.ApplyTransform(Transform, face.GetBox())).Center;
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

        public void SetParameter(string property, object value)
        {
            for (int i = 0; i < EquationManager.GetCount(); ++i)
            {
                string equation = EquationManager.Equation[i];
                if (equation.Contains(property))
                {
                    string newEquation = new Regex(@"=\s*(\d*)").Replace(equation, (m) => { return string.Format("={0}", value.ToString()); }, 1);
                    EquationManager.Delete(i);
                    EquationManager.Add2(i, newEquation, true);
                    Debug.WriteLine("replace equation {0} on {1} ... done!", equation, newEquation);
                    return;
                }
            }
        }

        public string GetParameter(string property)
        {
            for (int i = 0; i < EquationManager.GetCount(); ++i)
            {
                string equation = EquationManager.Equation[i];
                if (equation.Contains(property))
                {
                    Match match = new Regex(@"=\s*(\d*)").Match(equation);
                    if (match.Success)
                    {
                        Debug.WriteLine("get value {0} for property {1} ... done!", match.Groups[1].Value, property);
                        return match.Groups[1].Value;
                    }
                }
            }

            return null;
        }
    }
}
