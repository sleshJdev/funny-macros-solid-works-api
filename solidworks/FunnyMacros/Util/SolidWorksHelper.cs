using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using FunnyMacros.Model;

namespace FunnyMacros.Util
{
    public partial class SolidWorksMacro
    {
        //quantity of errors
        private int qe;

        //quantity of warnings
        private int qw;

        //paths to components
        private const string HOME_PATH = @"d:\!current-tasks\course-projects";
        private const string ASSEMBLY_NAME = "assembly.sldasm";
        private const string CORPUS_NAME = "corpus.sldprt";
        private const string SHAFT_NAME = "shaft.sldprt";
        private const string PLACE_LOCATOR_NAME = "plane-locator.sldprt";
        private const string CYLINDER_LOCATOR_NAME = "cylinder-locator.sldprt";

        //solidworks application instance
        public ISldWorks solidWorks;
        public IMathUtility mathUtility;

        //the assembly of details
        private IModelDoc2 document;
        private IAssemblyDoc assembly;

        //the corpus - the base for the design
        private Locator corpus = new Locator();

        //the detail for positionign
        private Locator shaft = new Locator();

        //bases of that detail
        private Base planeBase;
        private Base firstCylinderBase;
        private Base secondCylinderBase;

        //locator for bases
        private Locator planeLocator = new Locator();
        private Locator firstCylinderLocator = new Locator();
        private Locator secondCylinderLocator = new Locator();

        public void Open()
        {
            solidWorks = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("SldWorks.Application")) as ISldWorks;
            solidWorks.Visible = true;
        }

        public void Initalize()
        {
            document = LoadUtil.LoadAssembly(solidWorks, Path.Combine(HOME_PATH, ASSEMBLY_NAME));
            assembly = document as IAssemblyDoc;

            //load models
            corpus.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CORPUS_NAME));
            planeLocator.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, PLACE_LOCATOR_NAME));
            firstCylinderLocator.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME));
            secondCylinderLocator.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME));
            shaft.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, SHAFT_NAME));

            //adding to assemply
            corpus.Component = LoadUtil.AddModelToAssembly(assembly, corpus.Model.GetPathName(), 0.0, -60.0 / 1000, 0.0);
            planeLocator.Component = LoadUtil.AddModelToAssembly(assembly, planeLocator.Model.GetPathName(), 250.0 / 1000, 0.0, 0.0);
            firstCylinderLocator.Component = LoadUtil.AddModelToAssembly(assembly, firstCylinderLocator.Model.GetPathName(), 0.0, 0.0, 100.0 / 1000);
            secondCylinderLocator.Component = LoadUtil.AddModelToAssembly(assembly, secondCylinderLocator.Model.GetPathName(), 0.0, 0.0, -100.0 / 1000);
            shaft.Component = LoadUtil.AddModelToAssembly(assembly, shaft.Model.GetPathName(), 0.0, 0.0, 0.0);

            //fix shaft
            shaft.Component.Select4(true, null, false);
            assembly.FixComponent();
            document.ClearSelection2(true);

            //unfix locator 
            corpus.Component.Select4(true, null, false);
            planeLocator.Component.Select4(true, null, false);
            firstCylinderLocator.Component.Select4(true, null, false);
            secondCylinderLocator.Component.Select4(true, null, false);
            assembly.UnfixComponent();
            document.ClearSelection2(true);

            mathUtility = solidWorks.IGetMathUtility();
        }

        public void Commit()
        {
            document.EditRebuild3();
            //document.Rebuild((int)swRebuildOptions_e.swRebuildAll);//to apply translate
        }

        public void AlignWithHorizontAll()
        {
            //conjugation with the horizontal plane
            AlignWithHorizont(corpus.Component);
            AlignWithHorizont(shaft.Component);
            AlignWithHorizont(planeLocator.Component);
            AlignWithHorizont(firstCylinderLocator.Component);
            AlignWithHorizont(secondCylinderLocator.Component);
        }

        public void AlignWithHorizont(IComponent2 component)
        {
            document.ClearSelection2(true);
            IFeature face = component.FeatureByName("Сверху");
            IFeature horizont = assembly.IFeatureByName("Сверху");

            face = face == null ? component.FeatureByName("Top") : face;
            horizont = horizont == null ? assembly.IFeatureByName("Top") : horizont;

            face.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            horizont.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);

            int a = (int)swMateType_e.swMatePARALLEL;
            int b = (int)swMateAlign_e.swAlignNONE;
            assembly.AddMate3(a, b, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            document.ClearSelection2(true);
        }

        public void AddMate()
        {
            //mate for plane base
            IFeature feature = planeLocator.Component.FeatureByName("mouting-plane");
            AddMate(feature, planeBase.Face, (int)swMateType_e.swMateCOINCIDENT);

            double[] planeBaseCenter = MiscUtil.GetCenterOf(planeBase.Face.GetBox() as double[]);
            Debug.WriteLine("center of plane base: {0}", string.Join(" | ", planeBaseCenter), string.Empty);

            double[] mountPlaneCenter = MiscUtil.GetCenterOf(feature);
            Debug.WriteLine("center of mount plane: {0}", string.Join(" | ", planeBaseCenter), string.Empty);

            MiscUtil.Translate(mathUtility, planeLocator.Component,
                planeBaseCenter[0] - mountPlaneCenter[0]/*dx*/,
                planeBaseCenter[1] - mountPlaneCenter[1]/*dy*/,
                planeBaseCenter[2] - mountPlaneCenter[2]/*dz*/);

            Debug.WriteLine("mate with cylinder bases...");
            AddMateOfCylinderLocators(firstCylinderLocator.Component, firstCylinderBase.Face);
            AddMateOfCylinderLocators(secondCylinderLocator.Component, secondCylinderBase.Face);

            MiscUtil.Translate(mathUtility, firstCylinderLocator.Component, -0.3, -0.3, -0.3);
            MiscUtil.Translate(mathUtility, secondCylinderLocator.Component, 0.5, 0.5, 0.5);

            Commit();
            Debug.WriteLine("added mate ... ok");

            AlignWithShaft();
            Debug.WriteLine("align with shaft ... ok");
        }

        public void AddMateOfCylinderLocators(IComponent2 cylinderLocator, IFace2 cylinderBase)
        {
            IFeature firstFeature = cylinderLocator.FeatureByName("mount-plane-1");
            IFeature secondFeature = cylinderLocator.FeatureByName("mount-plane-2");

            AddMate(firstFeature, cylinderBase, (int)swMateType_e.swMateTANGENT);
            AddMate(secondFeature, cylinderBase, (int)swMateType_e.swMateTANGENT);
        }

        public void AddMate(IFeature locator, IFace2 face, int mateType)
        {
            document.ClearSelection2(true);
            locator.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            (face as IEntity).Select4(true, null);
            assembly.AddMate3(mateType, (int)swMateAlign_e.swAlignSAME, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            document.ClearSelection2(true);
        }

        public void AlignWithShaft()
        {
            Debug.WriteLine("find ends of shaft...");
            RemovalPairlFaces ends = MiscUtil.FindMaximumRemovalPlanes(mathUtility,
                Array.ConvertAll(shaft.Component.IGetBody().GetFaces(), new Converter<object, IFace2>((o) => { return o as IFace2; })));

            double[] baseBox = firstCylinderBase.Face.GetBox() as double[];
            Bounding(document, baseBox);

            double[] normal = firstCylinderBase.Face.Normal as double[];
            


            (firstCylinderBase as IEntity).Select4(true, null);
            //(ends.To as IEntity).Select4(true, null);
        }


        public void Bounding(IModelDoc2 model, double[] boxFeature)
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

            // Draw points at each corner of bounding box
            sketchPoint[0] = (SketchPoint)sketchManager.CreatePoint(boxFeature[3], boxFeature[1], boxFeature[5]);
            sketchPoint[1] = (SketchPoint)sketchManager.CreatePoint(boxFeature[0], boxFeature[1], boxFeature[5]);
            sketchPoint[2] = (SketchPoint)sketchManager.CreatePoint(boxFeature[0], boxFeature[1], boxFeature[2]);
            sketchPoint[3] = (SketchPoint)sketchManager.CreatePoint(boxFeature[3], boxFeature[1], boxFeature[2]);
            sketchPoint[4] = (SketchPoint)sketchManager.CreatePoint(boxFeature[3], boxFeature[4], boxFeature[5]);
            sketchPoint[5] = (SketchPoint)sketchManager.CreatePoint(boxFeature[0], boxFeature[4], boxFeature[5]);
            sketchPoint[6] = (SketchPoint)sketchManager.CreatePoint(boxFeature[0], boxFeature[4], boxFeature[2]);
            sketchPoint[7] = (SketchPoint)sketchManager.CreatePoint(boxFeature[3], boxFeature[4], boxFeature[2]);

            // Now draw bounding box
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

        public void DoProcessSelection()
        {
            ISelectionMgr selector = document.SelectionManager;
            int quantity = selector.GetSelectedObjectCount();
            if (quantity != 3)
            {
                int a = (int)swMessageBoxIcon_e.swMbWarning;
                int b = (int)swMessageBoxBtn_e.swMbOk;
                solidWorks.SendMsgToUser2("Make sure, what you pick 3 face!", a, b);
            }
            Debug.WriteLine("selected {0} surface", quantity);
            for (int i = 1; i <= quantity; ++i)
            {
                try
                {
                    IComponent2 component = selector.GetSelectedObjectsComponent4(i, -1) as IComponent2;
                    IFace2 face = selector.GetSelectedObject6(i, -1) as IFace2;
                    IFeature feature = face.IGetFeature();
                    ISurface surface = face.IGetSurface();
                    Debug.WriteLine("face {0}. details: materialIdName: {1}, materialUserName: {2}", i, face.MaterialIdName, face.MaterialUserName);
                    Debug.WriteLine("\tfeature details. name: {0}, visible: {1}, description: {2}", feature.Name, feature.Visible, feature.Description);
                    Debug.WriteLine("\tsurface details. isPlane: {0}, isCylinder: {1}", surface.IsPlane(), surface.IsCylinder());
                    if (surface.IsPlane() && planeBase == null)
                    {
                        planeBase.Face = face;
                        planeBase.Component = component;
                    }
                    else if (surface.IsCylinder())
                    {
                        if (firstCylinderBase == null)
                        {
                            firstCylinderBase.Face = face;
                            firstCylinderBase.Component = component;
                        }
                        else if(secondCylinderBase == null)
                        {
                            secondCylinderBase.Face = face;
                            secondCylinderBase.Component = component;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("process selected faces error: {0}. \n\tstack trace: {1}", ex.Message, ex.StackTrace);
                };
            }
            document.ClearSelection2(true);
            Debug.WriteLine("processing select of user ... ok");
        }

        public void Close()
        {
            if (solidWorks != null)
            {
                solidWorks.ExitApp();
            }
            Process[] processes = Process.GetProcessesByName("SLDWORKS");
            foreach (var process in processes)
            {
                try
                {
                    process.CloseMainWindow();
                    process.Kill();
                }
                catch (Exception e)
                {
                    Debug.WriteLine("error during close solidworks processes: {0}\nstack trace: {1}", e.Message, e.StackTrace);
                }
            }
        }
    }
}
