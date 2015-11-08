using System;

namespace FunnyMacros.Model
{
    class Box
    {
        public Box() : this(new double[6])
        {
        }

        public Box(double[] coordinatesCorners)
        {
            if (coordinatesCorners == null)
            {
                throw new ArgumentNullException("coordinatesCorners");
            }

            if (coordinatesCorners.Length != 6)
            {
                throw new ArgumentException("coordinatesCorners must have length 6", "coordinatesCorners");
            }

            CoordinatesCorners = coordinatesCorners;
        }

        public double[] CoordinatesCorners { get; set; }

        public double Volume
        {
            get { return Math.Abs(Xmax - Xmin) * Math.Abs(Ymax - Ymin) * Math.Abs(Zmax - Zmin); }
        }

        public double Xmin
        {
            get { return CoordinatesCorners[0]; }
            set { CoordinatesCorners[0] = value; }
        }

        public double Ymin
        {
            get { return CoordinatesCorners[1]; }
            set { CoordinatesCorners[1] = value; }
        }
        public double Zmin
        {
            get { return CoordinatesCorners[2]; }
            set { CoordinatesCorners[2] = value; }
        }

        public double Xmax
        {
            get { return CoordinatesCorners[3]; }
            set { CoordinatesCorners[3] = value; }
        }

        public double Ymax
        {
            get { return CoordinatesCorners[4]; }
            set { CoordinatesCorners[4] = value; }
        }

        public double Zmax
        {
            get { return CoordinatesCorners[5]; }
            set { CoordinatesCorners[5] = value; }
        }

        public override string ToString()
        {
            return string.Format("Box <{0}>, volume: {1}", string.Join(" | ", CoordinatesCorners), Volume);
        }
    }
}
