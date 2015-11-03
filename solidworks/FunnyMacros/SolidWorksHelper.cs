using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Threading;

namespace FunnyMacros
{
    class SolidWorksHelper
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

        //the assembly of details
        private IModelDoc2 document;
        private IAssemblyDoc assembly;

        //the corpus - the base for the design
        private Locator corpus = new Locator();

        //the detail for positionign
        private Locator shaft = new Locator();

        //bases of that detail
        private IFace2 planeBase;
        private IFace2 firstCylinderBase;
        private IFace2 secondCylinderBase;

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

            document = LoadAssembly(solidWorks, Path.Combine(HOME_PATH, ASSEMBLY_NAME));
            assembly = document as IAssemblyDoc;

            corpus.Model = LoadModel(Path.Combine(HOME_PATH, CORPUS_NAME));
            planeLocator.Model = LoadModel(Path.Combine(HOME_PATH, PLACE_LOCATOR_NAME));
            firstCylinderLocator.Model = LoadModel(Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME));
            secondCylinderLocator.Model = LoadModel(Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME));
            shaft.Model = LoadModel(Path.Combine(HOME_PATH, SHAFT_NAME));

            corpus.Component = AddModelToAssembly(assembly, corpus.Model.GetPathName(), 0.0, -50.0 / 1000, 0.0);
            planeLocator.Component = AddModelToAssembly(assembly, planeLocator.Model.GetPathName(), 250.0 / 1000, 0.0, 0.0);
            firstCylinderLocator.Component = AddModelToAssembly(assembly, firstCylinderLocator.Model.GetPathName(), 0.0, 0.0, 100.0 / 1000);
            secondCylinderLocator.Component = AddModelToAssembly(assembly, secondCylinderLocator.Model.GetPathName(), 0.0, 0.0, -100.0 / 1000);
            shaft.Component = AddModelToAssembly(assembly, shaft.Model.GetPathName(), 0.0, 0.0, 0.0);

            //fix shaft
            shaft.Component.Select(true);
            assembly.FixComponent();
            document.ClearSelection2(true);

            //unfix locator 
            corpus.Component.Select4(true, null, false);
            planeLocator.Component.Select4(true, null, false);
            firstCylinderLocator.Component.Select4(true, null, false);
            secondCylinderLocator.Component.Select4(true, null, false);
            assembly.UnfixComponent();
            document.ClearSelection2(true);
        }

        public IModelDoc2 LoadAssembly(ISldWorks application, string path)
        {
            //2 = (int)swDocTemplateTypes_e.swDocTemplateTypePART
            IModelDoc2 model = application.OpenDoc6(path, 2, 0, string.Empty, ref qe, ref qw);
            Debug.WriteLine("assemly loaded. errors: {2}, warnings: {3}, pathName: {0}, title: {1} ... ok", model.GetPathName(), model.GetTitle(), qe, qw);

            application.ActivateDoc2(model.GetTitle(), false, ref qe);
            Debug.WriteLine("activate doc. errors: {0} ... ok", qe, string.Empty);

            model = application.ActiveDoc as IModelDoc2;
            Debug.WriteLine("set assembly to active document ... ok");

            return model;
        }

        public IModelDoc2 LoadModel(string path)
        {
            //1  = (int)swDocTemplateTypes_e.swDocTemplateTypeNONE
            //16 = (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel
            IModelDoc2 model = solidWorks.OpenDoc6(path, 1, 16, string.Empty, ref qe, ref qw);
            Debug.WriteLine("model {0} loaded in mememory, full path {1} ... ok", model.GetTitle(), model.GetPathName());

            return model;
        }

        public IComponent2 AddModelToAssembly(IAssemblyDoc assembly, string componentName, double x, double y, double z)
        {
            int a = (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig;
            IComponent2 component = assembly.AddComponent5(componentName, a, string.Empty, false, string.Empty, x, y, z);
            Debug.WriteLine("model {0} added to assembly ... ok", componentName, string.Empty);

            return component;
        }

        public void AddMate()
        {
            //mate for plane base
            IFeature feature = planeLocator.Component.FeatureByName("mouting-plane");
            AddMate(feature, planeBase, (int)swMateType_e.swMateCOINCIDENT);

            double[] planeBaseCenter = GetCenterOf(planeBase.GetBox() as double[]);
            Debug.WriteLine("center of plane base: {0}", string.Join(" | ", planeBaseCenter), string.Empty);

            double[] mountPlaneCenter = GetCenterOf(feature);
            Debug.WriteLine("center of mount plane: {0}", string.Join(" | ", planeBaseCenter), string.Empty);

            Translate(planeLocator.Component, planeBaseCenter[0] - mountPlaneCenter[0], planeBaseCenter[1] - mountPlaneCenter[1], planeBaseCenter[2] - mountPlaneCenter[2]);

            IFeature firstFeature = null;
            IFeature secondFeature = null;

            Debug.WriteLine("mate with first cylinder base...");
            firstFeature = firstCylinderLocator.Component.FeatureByName("mount-plane-1");
            secondFeature = firstCylinderLocator.Component.FeatureByName("mount-plane-2");

            AddMate(firstFeature, firstCylinderBase, (int)swMateType_e.swMateTANGENT);
            AddMate(secondFeature, firstCylinderBase, (int)swMateType_e.swMateTANGENT);

            Translate(firstCylinderLocator.Component, -0.3, -0.3, -0.3);

            Debug.WriteLine("mate with second cylinder base...");
            firstFeature = secondCylinderLocator.Component.FeatureByName("mount-plane-1");
            secondFeature = secondCylinderLocator.Component.FeatureByName("mount-plane-2");

            AddMate(firstFeature, secondCylinderBase, (int)swMateType_e.swMateTANGENT);
            AddMate(secondFeature, secondCylinderBase, (int)swMateType_e.swMateTANGENT);

            Translate(secondCylinderLocator.Component, 0.5, 0.5, 0.5);

            Debug.WriteLine("added mate ... ok");

            AlignWithShaft();
        }

        public void AddMate(IFeature locator, IFace2 face, int mateType)
        {
            locator.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            (face as IEntity).Select4(true, null);
            assembly.AddMate3(mateType, (int)swMateAlign_e.swAlignSAME, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            document.ClearSelection2(true);
        }

        public void AddMate(IFace2 end, IFace2 locatorFace, int mateType)
        {
            (end as IEntity).Select4(true, null);
            (locatorFace as IEntity).Select4(true, null);
            assembly.AddMate3(mateType, (int)swMateAlign_e.swAlignSAME, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            document.ClearSelection2(true);
        }

        public void AlignWithHorizong(IComponent2 component1, IComponent2 component2)
        {
            document.ClearSelection2(true);
            IFeature face1 = component1.FeatureByName("Сверху");
            Feature face2 = component2.FeatureByName("Сверху");

            face1 = face1 == null ? component1.FeatureByName("Top") : face1;
            face2 = face2 == null ? component2.FeatureByName("Top") : face2;

            face1.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            face2.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);

            int a = (int)swMateType_e.swMatePARALLEL;
            int b = (int)swMateAlign_e.swMateAlignALIGNED;
            assembly.AddMate3(a, b, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            document.ClearSelection2(true);
        }

        public void AlignWithShaft()
        {
            //conjugation with the horizontal plane
            AlignWithHorizong(planeLocator.Component, shaft.Component);
            AlignWithHorizong(firstCylinderLocator.Component, shaft.Component);
            AlignWithHorizong(secondCylinderLocator.Component, shaft.Component);

            Debug.WriteLine("iterate faces");

            IBody2 shaftBody = shaft.Component.GetBody();
            object[] faces = shaftBody.GetFaces();
            IFace2 firstEnd = null;
            IFace2 secondEnd = null;
            double minDistance = double.MinValue;
            double maxDistance = double.MaxValue;

            foreach (object o in faces)
            {                
                Face2 face = o as Face2;
                Debug.WriteLine("faceId: {0}, featureId: {1}", face.GetFaceId(), face.GetFeatureId());
                Surface surface = face.GetSurface();
                if (surface.IsPlane())
                {
                    double[] p = surface.PlaneParams;
                    double distance = Math.Sqrt(Math.Pow(p[0] * p[3], 2) + Math.Pow(p[1] * p[4], 2) + Math.Pow(p[2] * p[5], 2));
                    if (distance < maxDistance)
                    {
                        maxDistance = distance;
                        firstEnd = face;
                    }
                    else
                    {
                        minDistance = distance;
                        secondEnd = face;
                    }

                    //Debug.WriteLine(string.Join(" | ", p));
                    //Thread.Sleep(2000);
                }
            }

            (firstEnd as IEntity).Select4(true, null);
            (secondEnd as IEntity).Select4(true, null);

            IFace2 firstNearestFace = FindRemoteFaceOfCylinderLocator(firstCylinderLocator, firstEnd);
            //AddMate(firstEnd, firstNearestFace, (int)swMateType_e.swMateCOINCIDENT);            
            (firstNearestFace as IEntity).Select4(true, null);
        }

        public IFace2 FindRemoteFaceOfCylinderLocator(Locator locator, IFace2 originFace)
        {
            Debug.WriteLine("FindRemoteFaceOfCylinderLocator");
            double[] p = (originFace.GetSurface() as ISurface).PlaneParams;
            Debug.WriteLine("origin     " + string.Join(" | ", p));
            IBody2 body = locator.Component.GetBody();
            object[] faces = body.GetFaces();
            IFace2 f = null;
            
            foreach (object o in faces)
            {
                IFace2 face = o as IFace2;
                (face as IEntity).Select4(true, null);
                ISurface surface = face.GetSurface();
                if (surface.IsPlane())
                {
                    double[] q = surface.PlaneParams;
                    Debug.WriteLine(string.Join(" | ", q));
                    if (p[0] == q[0] && p[1] == q[1] && p[2] == q[2])
                    {
                        f = face;
                    }
                    Thread.Sleep(1000);
                }
            }

            return f;
        }

        public void Translate(IComponent2 component, double dx, double dy, double dz)
        {
            IMathUtility mathUtil = solidWorks.GetMathUtility();
            IMathTransform transform = mathUtil.CreateTransform(null);
            double[] matrix = transform.ArrayData as double[];
            matrix[9] = dx;
            matrix[10] = dy;
            matrix[11] = dz;
            component.Transform2 = mathUtil.CreateTransform(matrix);
            document.EditRebuild3();
        }

        public double[] GetCenterOf(IFeature feature)
        {
            object box = null;
            feature.GetBox(ref box);

            return GetCenterOf(box as double[]);
        }

        public double[] GetCenterOf(double[] coordinatesCorner)
        {
            Debug.WriteLine("calculate center of coordinates: {0}", string.Join(" | ", coordinatesCorner), "");

            double[] center = new double[3];
            for (int i = 0; i < center.Length; ++i)
            {
                center[i] = (coordinatesCorner[i] + coordinatesCorner[i + 3]) / 2;
            }

            return center;
        }

        public void DoProcessSelection()
        {
            ISelectionMgr selector = document.SelectionManager;
            int quantity = selector.GetSelectedObjectCount();
            //if (quantity != 3)
            //{
            //    throw new ArgumentException("Make sure, what you pick 3 face!");
            //}
            Debug.WriteLine("selected {0} surface", quantity);
            for (int i = 1; i <= quantity; ++i)
            {
                try
                {
                    IFace2 face = (IFace2)selector.GetSelectedObject6(i, -1);
                    IFeature feature = (IFeature)face.GetFeature();
                    ISurface surface = (ISurface)face.GetSurface();
                    Debug.WriteLine("face {0}. details: materialIdName: {1}, materialUserName: {2}", i, face.MaterialIdName, face.MaterialUserName);
                    Debug.WriteLine("\tfeature details. name: {0}, visible: {1}, description: {2}", feature.Name, feature.Visible, feature.Description);
                    Debug.WriteLine("\tsurface details. isPlane: {0}, isCylinder: {1}", surface.IsPlane(), surface.IsCylinder());
                    planeBase = planeBase == null && surface.IsPlane() ? face : planeBase;
                    firstCylinderBase = firstCylinderBase == null && surface.IsCylinder() ? face : firstCylinderBase;
                    secondCylinderBase = secondCylinderBase == null && surface.IsCylinder() ? face : secondCylinderBase;
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

        public void Contour(IModelDoc2 model)
        {
            const double MaxDouble = 1.79769313486231E+308;
            const double MinDouble = -1.79769313486231E+308;
            IModelDoc2 swModel = default(IModelDoc2);
            IAssemblyDoc swAssy = default(IAssemblyDoc);
            IConfiguration swConfig = default(IConfiguration);
            IConfigurationManager swConfigurationMgr = default(IConfigurationManager);
            IComponent2 swRootComp = default(IComponent2);
            object[] vChild = null;
            IComponent2 swChildComp = default(IComponent2);
            object Box = null;
            double[] BoxArray = new double[6];
            double X_max = 0;
            double X_min = 0;
            double Y_max = 0;
            double Y_min = 0;
            double Z_max = 0;
            double Z_min = 0;
            SketchManager swSketchMgr = default(SketchManager);
            SketchPoint[] swSketchPt = new SketchPoint[9];
            SketchSegment[] swSketchSeg = new SketchSegment[13];
            int i = 0;

            swModel = model;
            swAssy = (IAssemblyDoc)swModel;
            swConfigurationMgr = (ConfigurationManager)swModel.ConfigurationManager;
            swConfig = (Configuration)swConfigurationMgr.ActiveConfiguration;
            swRootComp = (Component2)swConfig.GetRootComponent3(true);


            // Initialize values
            X_max = MinDouble;
            X_min = MaxDouble;
            Y_max = MinDouble;
            Y_min = MaxDouble;
            Z_max = MinDouble;
            Z_min = MaxDouble;

            vChild = (object[])swRootComp.GetChildren();
            for (i = 0; i <= vChild.GetUpperBound(0); i++)
            {
                swChildComp = (Component2)vChild[i];
                if (swChildComp.Visible == (int)swComponentVisibilityState_e.swComponentVisible)
                {
                    Box = (object)swChildComp.GetBox(false, false);
                    BoxArray = (double[])Box;
                    X_max = GetMax(BoxArray[0], BoxArray[3], X_max);
                    X_min = GetMin(BoxArray[0], BoxArray[3], X_min);
                    Y_max = GetMax(BoxArray[1], BoxArray[4], Y_max);
                    Y_min = GetMin(BoxArray[1], BoxArray[4], Y_min);
                    Z_max = GetMax(BoxArray[2], BoxArray[5], Z_max);
                    Z_min = GetMin(BoxArray[2], BoxArray[5], Z_min);
                }
            }

            Debug.Print("Assembly Bounding Box (" + swModel.GetPathName() + ") = ");
            Debug.Print("  (" + (X_min * 1000.0) + "," + (Y_min * 1000.0) + "," + (Z_min * 1000.0) + ") mm");
            Debug.Print("  (" + (X_max * 1000.0) + "," + (Y_max * 1000.0) + "," + (Z_max * 1000.0) + ") mm");

            swSketchMgr = swModel.SketchManager;
            swSketchMgr.Insert3DSketch(true);
            swSketchMgr.AddToDB = true;



            // Draw points at each corner of bounding box
            swSketchPt[0] = swSketchMgr.CreatePoint(X_min, Y_min, Z_min);
            swSketchPt[1] = swSketchMgr.CreatePoint(X_min, Y_min, Z_max);
            swSketchPt[2] = swSketchMgr.CreatePoint(X_min, Y_max, Z_min);
            swSketchPt[3] = swSketchMgr.CreatePoint(X_min, Y_max, Z_max);
            swSketchPt[4] = swSketchMgr.CreatePoint(X_max, Y_min, Z_min);
            swSketchPt[5] = swSketchMgr.CreatePoint(X_max, Y_min, Z_max);
            swSketchPt[6] = swSketchMgr.CreatePoint(X_max, Y_max, Z_min);
            swSketchPt[7] = swSketchMgr.CreatePoint(X_max, Y_max, Z_max);

            // Draw bounding box
            swSketchSeg[0] = swSketchMgr.CreateLine(X_min, Y_min, Z_min, X_max, Y_min, Z_min);
            swSketchSeg[1] = swSketchMgr.CreateLine(X_max, Y_min, Z_min, X_max, Y_min, Z_max);
            swSketchSeg[2] = swSketchMgr.CreateLine(X_max, Y_min, Z_max, X_min, Y_min, Z_max);
            swSketchSeg[3] = swSketchMgr.CreateLine(X_min, Y_min, Z_max, X_min, Y_min, Z_min);
            swSketchSeg[4] = swSketchMgr.CreateLine(X_min, Y_min, Z_min, X_min, Y_max, Z_min);
            swSketchSeg[5] = swSketchMgr.CreateLine(X_min, Y_min, Z_max, X_min, Y_max, Z_max);
            swSketchSeg[6] = swSketchMgr.CreateLine(X_max, Y_min, Z_min, X_max, Y_max, Z_min);
            swSketchSeg[7] = swSketchMgr.CreateLine(X_max, Y_min, Z_max, X_max, Y_max, Z_max);
            swSketchSeg[8] = swSketchMgr.CreateLine(X_min, Y_max, Z_min, X_max, Y_max, Z_min);
            swSketchSeg[9] = swSketchMgr.CreateLine(X_max, Y_max, Z_min, X_max, Y_max, Z_max);
            swSketchSeg[10] = swSketchMgr.CreateLine(X_max, Y_max, Z_max, X_min, Y_max, Z_max);
            swSketchSeg[11] = swSketchMgr.CreateLine(X_min, Y_max, Z_max, X_min, Y_max, Z_min);

            swSketchMgr.AddToDB = false;
            swSketchMgr.Insert3DSketch(true);
        }

        public double GetMax(double Val1, double Val2, double Val3)
        {
            double functionReturnValue = 0;
            // Finds maximum of 3 values
            functionReturnValue = Val1;
            if (Val2 > functionReturnValue)
            {
                functionReturnValue = Val2;
            }
            if (Val3 > functionReturnValue)
            {
                functionReturnValue = Val3;
            }
            return functionReturnValue;
        }
        public double GetMin(double Val1, double Val2, double Val3)
        {
            double functionReturnValue = 0;
            // Finds minimum of 3 values
            functionReturnValue = Val1;
            if (Val2 < functionReturnValue)
            {
                functionReturnValue = Val2;
            }
            if (Val3 < functionReturnValue)
            {
                functionReturnValue = Val3;
            }
            return functionReturnValue;
        }
    }
}
