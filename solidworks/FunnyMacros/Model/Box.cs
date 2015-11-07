using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FunnyMacros.Model
{
    class Box
    {
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

            Xmin = coordinatesCorners[0];
            Ymin = coordinatesCorners[1];
            Zmin = coordinatesCorners[2];

            Xmax = coordinatesCorners[3];
            Ymax = coordinatesCorners[4];
            Zmax = coordinatesCorners[5];
        }

        public double[] CoordinatesCorners { get; set; }

        public double Xmin { get; set; }
        public double Ymin { get; set; }
        public double Zmin { get; set; }

        public double Xmax { get; set; }
        public double Ymax { get; set; }
        public double Zmax { get; set; }
    }
}
