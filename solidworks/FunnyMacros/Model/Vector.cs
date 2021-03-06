﻿using System;

namespace FunnyMacros.Model
{
    class Vector
    {
        public Vector() : this(0.0, 0.0, 0.0)
        {
        }

        public Vector(double x, double y, double z) 
        {
            X = x;
            Y = y;
            Z = z;
        }

        public Vector(double[] coordinates)
        {
            if (coordinates == null)
            {
                throw new ArgumentNullException("coordinates");
            }

            if (coordinates.Length != 3)
            {
                throw new ArgumentException("coordinates must have length 6", "coordinates");
            }

            Coordinates = coordinates;

            X = coordinates[0];
            Y = coordinates[1];
            Z = coordinates[2];
        }

        public double[] Coordinates { get; set; }

        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }

        public override string ToString()
        {
            return string.Format("Vector <{0}>", string.Join(" | ", Coordinates), string.Empty);
        }
    }
}
