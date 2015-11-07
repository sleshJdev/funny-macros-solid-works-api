using System.Diagnostics;
using SolidWorks.Interop.sldworks;

namespace FunnyMacros.Util
{
    class Drawer
    {
        public static void DrawPoint(IModelDoc2 model, double x, double y, double z)
        {
            Debug.WriteLine("draw point: [{0}, {1}, {2}]", x, y, z);
            model.Insert3DSketch2(true);
            model.SetAddToDB(true);
            model.SetDisplayWhenAdded(false);
            model.SketchManager.CreatePoint(x, y, z);
            model.SetDisplayWhenAdded(true);
            model.SetAddToDB(false);
            model.Insert3DSketch2(true);
        }

        public static void DrawLine(IModelDoc2 model, double x1, double y1, double z1, double x2, double y2, double z2)
        {
            model.Insert3DSketch2(true);
            model.SetAddToDB(true);
            model.SetDisplayWhenAdded(false);
            ISketchPoint from = model.SketchManager.CreatePoint(x1, y1, z1);
            ISketchPoint to = model.SketchManager.CreatePoint(x2, y2, z2);
            ISketchSegment line = model.SketchManager.CreateLine(from.X, from.Y, from.Z, to.X, to.Y, to.Z);
            Debug.WriteLine("draw line(length = {6}): from [{0}, {1}, {2}] to [{3}, {4}, {5}]", x1, y1, z1, x2, y2, z2, line.GetLength());
            model.SketchManager.CreatePoint(x2, y2, z2);
            model.SetDisplayWhenAdded(true);
            model.SetAddToDB(false);
            model.Insert3DSketch2(true);
        }

        public static void Bounding(IModelDoc2 model, double[] boxFeature)
        {
            SketchManager sketchManager = default(SketchManager);
            SketchPoint[] sketchPoint = new SketchPoint[9];
            SketchSegment[] sketchSegment = new SketchSegment[13];

            Debug.Print("  point1 = " + "(" + boxFeature[0] * 1000.0 + ", " + boxFeature[1] * 1000.0 + ", " + boxFeature[2] * 1000.0 + ") mm");
            Debug.Print("  point2 = " + "(" + boxFeature[3] * 1000.0 + ", " + boxFeature[4] * 1000.0 + ", " + boxFeature[5] * 1000.0 + ") mm");
            model.Insert3DSketch2(true);
            model.SetAddToDB(true);
            model.SetDisplayWhenAdded(false);

            sketchManager = (SketchManager)model.SketchManager;

            // draw points at each corner of bounding box
            sketchPoint[0] = (SketchPoint)sketchManager.CreatePoint(boxFeature[3], boxFeature[1], boxFeature[5]);
            sketchPoint[1] = (SketchPoint)sketchManager.CreatePoint(boxFeature[0], boxFeature[1], boxFeature[5]);
            sketchPoint[2] = (SketchPoint)sketchManager.CreatePoint(boxFeature[0], boxFeature[1], boxFeature[2]);
            sketchPoint[3] = (SketchPoint)sketchManager.CreatePoint(boxFeature[3], boxFeature[1], boxFeature[2]);
            sketchPoint[4] = (SketchPoint)sketchManager.CreatePoint(boxFeature[3], boxFeature[4], boxFeature[5]);
            sketchPoint[5] = (SketchPoint)sketchManager.CreatePoint(boxFeature[0], boxFeature[4], boxFeature[5]);
            sketchPoint[6] = (SketchPoint)sketchManager.CreatePoint(boxFeature[0], boxFeature[4], boxFeature[2]);
            sketchPoint[7] = (SketchPoint)sketchManager.CreatePoint(boxFeature[3], boxFeature[4], boxFeature[2]);

            // now draw bounding box
            sketchSegment[0] = (SketchSegment)sketchManager.CreateLine(sketchPoint[0].X, sketchPoint[0].Y, sketchPoint[0].Z, sketchPoint[1].X, sketchPoint[1].Y, sketchPoint[1].Z);
            sketchSegment[1] = (SketchSegment)sketchManager.CreateLine(sketchPoint[1].X, sketchPoint[1].Y, sketchPoint[1].Z, sketchPoint[2].X, sketchPoint[2].Y, sketchPoint[2].Z);
            sketchSegment[2] = (SketchSegment)sketchManager.CreateLine(sketchPoint[2].X, sketchPoint[2].Y, sketchPoint[2].Z, sketchPoint[3].X, sketchPoint[3].Y, sketchPoint[3].Z);
            sketchSegment[3] = (SketchSegment)sketchManager.CreateLine(sketchPoint[3].X, sketchPoint[3].Y, sketchPoint[3].Z, sketchPoint[0].X, sketchPoint[0].Y, sketchPoint[0].Z);
            sketchSegment[4] = (SketchSegment)sketchManager.CreateLine(sketchPoint[0].X, sketchPoint[0].Y, sketchPoint[0].Z, sketchPoint[4].X, sketchPoint[4].Y, sketchPoint[4].Z);
            sketchSegment[5] = (SketchSegment)sketchManager.CreateLine(sketchPoint[1].X, sketchPoint[1].Y, sketchPoint[1].Z, sketchPoint[5].X, sketchPoint[5].Y, sketchPoint[5].Z);
            sketchSegment[6] = (SketchSegment)sketchManager.CreateLine(sketchPoint[2].X, sketchPoint[2].Y, sketchPoint[2].Z, sketchPoint[6].X, sketchPoint[6].Y, sketchPoint[6].Z);
            sketchSegment[7] = (SketchSegment)sketchManager.CreateLine(sketchPoint[3].X, sketchPoint[3].Y, sketchPoint[3].Z, sketchPoint[7].X, sketchPoint[7].Y, sketchPoint[7].Z);
            sketchSegment[8] = (SketchSegment)sketchManager.CreateLine(sketchPoint[4].X, sketchPoint[4].Y, sketchPoint[4].Z, sketchPoint[5].X, sketchPoint[5].Y, sketchPoint[5].Z);
            sketchSegment[9] = (SketchSegment)sketchManager.CreateLine(sketchPoint[5].X, sketchPoint[5].Y, sketchPoint[5].Z, sketchPoint[6].X, sketchPoint[6].Y, sketchPoint[6].Z);
            sketchSegment[10] = (SketchSegment)sketchManager.CreateLine(sketchPoint[6].X, sketchPoint[6].Y, sketchPoint[6].Z, sketchPoint[7].X, sketchPoint[7].Y, sketchPoint[7].Z);
            sketchSegment[11] = (SketchSegment)sketchManager.CreateLine(sketchPoint[7].X, sketchPoint[7].Y, sketchPoint[7].Z, sketchPoint[4].X, sketchPoint[4].Y, sketchPoint[4].Z);

            model.SetDisplayWhenAdded(true);
            model.SetAddToDB(false);
            model.Insert3DSketch2(true);
        }
    }
}
