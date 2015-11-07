using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using SolidWorks.Interop.sldworks;
using FunnyMacros.Model;

namespace FunnyMacros.Util
{
    class MiscUtil
    {
        public static double[] GetInGlobalCoordinatesSystem(IMathUtility mathUtility, IMathTransform transform, double[] vector)
        {
            IMathPoint point = null;
            double[] coordinates = new double[3];
            double[] result = new double[vector.Length];
            for (int i = 0; i < vector.Length; i += 3)
            {
                coordinates[0] = vector[i + 0];
                coordinates[1] = vector[i + 1];
                coordinates[2] = vector[i + 2];
                point = mathUtility.CreatePoint(coordinates);
                point = point.MultiplyTransform(transform);
                coordinates = point.ArrayData as double[];
                result[i + 0] = coordinates[0];
                result[i + 1] = coordinates[1];
                result[i + 2] = coordinates[2];
            }

            return result;
        }

        public static double[] GetCenterBoxOf(IFeature feature)
        {
            object box = null;
            feature.GetBox(ref box);

            return GetCenterOf(box as double[]);
        }

        public static double[] GetCenterOf(double[] coordinatesCorner)
        {
            Debug.WriteLine("calculate center of coordinates: {0}", string.Join(" | ", coordinatesCorner), string.Empty);
            return new double[]
            {
                (coordinatesCorner[0] + coordinatesCorner[3]) / 2,
                (coordinatesCorner[1] + coordinatesCorner[4]) / 2,
                (coordinatesCorner[2] + coordinatesCorner[5]) / 2
            };
        }

        public static void Translate(IMathUtility mathUtility, IComponent2 component, double x, double y, double z)
        {
            IMathTransform transform = mathUtility.CreateTransform(null);
            double[] matrix = transform.ArrayData as double[];
            matrix[9] = x;
            matrix[10] = y;
            matrix[11] = z;
            component.Transform2 = mathUtility.CreateTransform(matrix);
        }

        public static void FindRemovalFace(IMathUtility mathUtility, IFace2[] faces, IFace2 fromFace, out IFace2 removalFace, out double distance)
        {
            distance = 0.0;
            removalFace = null;
            ISurface fromSurface = fromFace.GetSurface();
            IMathVector fromFaceNormal = mathUtility.CreateVector(fromFace.Normal);

            for (int i = 0; i < faces.Length; ++i)
            {
                IFace2 toFace = faces[i];
                ISurface toSurface = toFace.GetSurface();
                if (toSurface.IsPlane())
                {
                    IMathVector toFaceNormal = mathUtility.CreateVector(toFace.Normal);
                    if (IsCoDirectional(fromFaceNormal, toFaceNormal) || IsCoDirectional(fromFaceNormal, Negative(mathUtility, toFaceNormal)))
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

        public static RemovalPairlFaces FindMaximumRemovalPlanes(IMathUtility mathUtility, IFace2[] faces)
        {
            List<RemovalPairlFaces> list = new List<RemovalPairlFaces>();
            for (int i = 0; i < faces.Length; ++i)
            {
                IFace2 fromFace = faces[i];
                if (fromFace.IGetSurface().IsPlane())
                {                    
                    IFace2 toFace;
                    double distance;
                    FindRemovalFace(mathUtility, faces.Skip(i + 1).ToArray(), fromFace, out toFace, out distance);

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
        public static double DistanceBetweenParallelPlanes(ISurface plane1, ISurface plane2)
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

        public static bool IsCoDirectional(IMathVector one, IMathVector two)
        {
            IMathVector oneN = one.Normalise();
            IMathVector twoN = two.Normalise();

            double[] oneVector = oneN.ArrayData as double[];
            double[] twoVector = twoN.ArrayData as double[];
            double[] difference = new double[3];

            difference[0] = oneVector[0] - twoVector[0];
            difference[1] = oneVector[1] - twoVector[1];
            difference[2] = oneVector[2] - twoVector[2];

            return Equal(difference[0]) && Equal(difference[1]) && Equal(difference[2]);
        }

        public static bool Equal(double delta)
        {
            return Math.Abs(delta) < 0.15/*precision*/;
        }

        public static IMathVector Negative(IMathUtility mathUtility, IMathVector origin)
        {
            double[] negativeVector = origin.ArrayData as double[];
            negativeVector[0] = -negativeVector[0];
            negativeVector[1] = -negativeVector[1];
            negativeVector[2] = -negativeVector[2];

            return mathUtility.CreateVector(negativeVector);
        }
    }
}
