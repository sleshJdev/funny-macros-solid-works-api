using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FunnyMacros
{
    class SolitUtil
    {
        public IMathUtility MathUtility { get; set; }

        public SolitUtil(IMathUtility mathUtility)
        {
            MathUtility = mathUtility;
        }

        public void FindRemovalFace(IFace2[] faces, IFace2 fromFace, out IFace2 removalFace, out double distance)
        {
            distance = 0.0;
            removalFace = null;
            ISurface fromSurface = fromFace.GetSurface();
            IMathVector fromFaceNormal = MathUtility.CreateVector(fromFace.Normal);

            for (int i = 0; i < faces.Length; ++i)
            {
                IFace2 toFace = faces[i];
                ISurface toSurface = toFace.GetSurface();
                if (toSurface.IsPlane())
                {
                    IMathVector toFaceNormal = MathUtility.CreateVector(toFace.Normal);
                    if (IsCoDirectional(fromFaceNormal, toFaceNormal) || IsCoDirectional(fromFaceNormal, Negative(toFaceNormal)))
                    {
                        double currentDistance = DistanceBetweenParallelPlanes(fromSurface, toSurface);
                        if (currentDistance > distance)
                        {
                            removalFace = toFace;
                            distance = currentDistance;
                        }
                    }                    
                }
            }
        }

        public RemovalPairlFaces FindMaximumRemovalPlanes(IFace2[] faces)
        {
            List<RemovalPairlFaces> list = new List<RemovalPairlFaces>();
            for (int i = 0; i < faces.Length; ++i)
            {
                IFace2 fromFace = faces[i];
                ISurface fromSurface = fromFace.GetSurface();
                if (fromSurface.IsPlane())
                {                    
                    IFace2 toFace;
                    double distance;
                    FindRemovalFace(faces, fromFace, out toFace, out distance);

                    if (toFace != null)
                    {
                        list.Add(new RemovalPairlFaces()
                        {
                            From = fromFace,
                            To = toFace,
                            Distance = distance
                        });
                    }
                    
                }
            }

            return list.OrderBy((x) => { return -x.Distance; }).First();
        }


        /// <summary>
        /// Calculate distance between parallel planes
        /// </summary>
        /// <param name="plane1"></param>
        /// <param name="plane2"></param>
        /// <returns></returns>
        public double DistanceBetweenParallelPlanes(ISurface plane1, ISurface plane2)
        {
            double[] params1 = plane1.PlaneParams as double[];
            double[] parans2 = plane2.PlaneParams as double[];
            double distance;

            //get equation for first plane
            //aX + bY + cZ + d = 0
            double a, b, c, d;
            a = params1[0];
            b = params1[1];
            c = params1[2];
            d = -a * params1[3] - b * params1[4] - c * params1[5];

            //calculate distance
            distance = Math.Abs(a * parans2[3] + b * parans2[4] + c * parans2[5] + d) / Math.Sqrt(Math.Pow(a, 2) + Math.Pow(b, 2) + Math.Pow(c, 2));

            return distance;
        }

        public bool IsCoDirectional(IMathVector one, IMathVector two)
        {
            IMathVector oneN = one.Normalise();
            IMathVector twoN = two.Normalise();

            double[] oneVector = oneN.ArrayData as double[];
            double[] twoVector = twoN.ArrayData as double[];
            double[] difference = new double[3];

            difference[0] = oneVector[0] - twoVector[0];
            difference[1] = oneVector[1] - twoVector[1];
            difference[2] = oneVector[2] - twoVector[2];

            return Equal(difference[0], 0, 1) && Equal(difference[1], 0, 1) && Equal(difference[2], 0, 1);
        }

        public bool Equal(double d1, double d2, int toleranceDecimalPlaces)
        {
            double tolerance = 1.0 / (Math.Pow(10, toleranceDecimalPlaces));

            return Math.Abs(d1 - d2) < tolerance;
        }

        public IMathVector Negative(IMathVector origin)
        {
            double[] negativeVector = origin.ArrayData as double[];
            negativeVector[0] = -negativeVector[0];
            negativeVector[1] = -negativeVector[1];
            negativeVector[2] = -negativeVector[2];

            return MathUtility.CreateVector(negativeVector);
        }
    }
}
