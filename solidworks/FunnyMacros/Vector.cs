using System;

namespace FunnyMacros
{
    class Vector
    {
        public double X { set; get; }
        public double Y { set; get; }
        public double Z { set; get; }

        public Vector()
        {}

        public Vector(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public double Distance(Vector v)
        {
            return Math.Sqrt(Math.Pow(X - v.X, 2) + Math.Pow(Y - v.Y, 2) + Math.Pow(Z - v.Z, 2));
        }
    }
}
