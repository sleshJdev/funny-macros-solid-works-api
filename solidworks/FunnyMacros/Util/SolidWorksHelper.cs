using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using FunnyMacros.Model;

namespace FunnyMacros.Util
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
            document = LoadUtil.LoadAssembly(solidWorks, Path.Combine(HOME_PATH, ASSEMBLY_NAME));
            assembly = document as IAssemblyDoc;

            //load models
            corpus.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CORPUS_NAME));
            planeLocator.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, PLACE_LOCATOR_NAME));
            firstCylinderLocator.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME));
            secondCylinderLocator.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME));
            shaft.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, SHAFT_NAME));

            //adding to assemply
            corpus.Component = LoadUtil.AddModelToAssembly(assembly, corpus.Model.GetPathName(), 0.0, -50.0 / 1000, 0.0);
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

            //conjugation with the horizontal plane
            AlignWithHorizong(corpus.Component);
            AlignWithHorizong(shaft.Component);
            AlignWithHorizong(planeLocator.Component);
            AlignWithHorizong(firstCylinderLocator.Component);
            AlignWithHorizong(secondCylinderLocator.Component);
        }

        public void AlignWithHorizong(IComponent2 component)
        {
            document.ClearSelection2(true);
            IFeature face = component.FeatureByName("Сверху");
            IFeature horizont = assembly.IFeatureByName("Сверху");

            face = face == null ? component.FeatureByName("Top") : face;
            horizont = horizont == null ? assembly.IFeatureByName("Top") : horizont;

            face.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            horizont.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);

            int a = (int)swMateType_e.swMatePARALLEL;
            int b = (int)swMateAlign_e.swMateAlignALIGNED;
            assembly.AddMate3(a, b, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            document.ClearSelection2(true);
        }

        public void AddMate()
        {
            //mate for plane base
            IFeature feature = planeLocator.Component.FeatureByName("mouting-plane");
            AddMate(feature, planeBase, (int)swMateType_e.swMateCOINCIDENT);

            double[] planeBaseCenter = MiscUtil.GetCenterOf(planeBase.GetBox() as double[]);
            Debug.WriteLine("center of plane base: {0}", string.Join(" | ", planeBaseCenter), string.Empty);

            double[] mountPlaneCenter = MiscUtil.GetCenterOf(feature);
            Debug.WriteLine("center of mount plane: {0}", string.Join(" | ", planeBaseCenter), string.Empty);

            MiscUtil.Translate(solidWorks.IGetMathUtility(), planeLocator.Component, 
                planeBaseCenter[0] - mountPlaneCenter[0]/*dx*/, 
                planeBaseCenter[1] - mountPlaneCenter[1]/*dy*/,
                planeBaseCenter[2] - mountPlaneCenter[2]/*dz*/);

            Debug.WriteLine("mate with cylinder bases...");
            AddMateOfCylinderLocators(firstCylinderLocator.Component, firstCylinderBase);
            AddMateOfCylinderLocators(secondCylinderLocator.Component, secondCylinderBase);

            MiscUtil.Translate(solidWorks.IGetMathUtility(), secondCylinderLocator.Component, 0.5, 0.5, 0.5);
            MiscUtil.Translate(solidWorks.IGetMathUtility(), firstCylinderLocator.Component, -0.3, -0.3, -0.3);

            document.EditRebuild3();//to apply translate

            Debug.WriteLine("added mate ... ok");
            AlignWithShaft();
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
            IBody2 body = null;
            object[] faceObjects = null;
            IFace2[] faces = null;

            Debug.WriteLine("find ends of shaft...");
            body = shaft.Component.GetBody();
            faceObjects = body.GetFaces();
            faces = Array.ConvertAll(faceObjects, new Converter<object, IFace2>((o) => { return o as IFace2; }));
            RemovalPairlFaces removalPair = MiscUtil.FindMaximumRemovalPlanes(solidWorks.IGetMathUtility(), faces);

            (removalPair.From as IEntity).Select4(true, null);
            (removalPair.To as IEntity).Select4(true, null);




            //IFace2 removalFace;
            //double distance;

            //body = firstCylinderLocator.Component.GetBody();
            //faceObjects = body.GetFaces();
            //faces = Array.ConvertAll(faceObjects, new Converter<object, IFace2>((o) => { return o as IFace2; }));
            //solidUtil.FindRemovalFace(faces, removalPair.From, out removalFace, out distance);

            //double[] normal1 = removalPair.From.Normal;
            //double[] normal2 = removalPair.To.Normal;
            //Debug.WriteLine("normal1: " + string.Join(" | ", normal1));
            //Debug.WriteLine("normal2: " + string.Join(" | ", normal2));

            //IMathVector v1 = solidUtil.MathUtility.CreateVector(normal1);

            //foreach (IFace2 face in faces)
            //{
            //    (face as Entity).Select4(true, null);
            //    Surface s = face.GetSurface();
            //    double[] ar = s.PlaneParams;
            //    double[] normal = face.Normal;
            //    Debug.WriteLine(string.Join(" | ", normal));
            //    IMathVector v = solidUtil.MathUtility.CreateVector(normal);
            //    IMathVector vnorm = v.Normalise();
            //    Debug.WriteLine("params: " + string.Join(" | ", ar));
            //    Debug.WriteLine("normalize vector: " + string.Join(" | ", (vnorm.ArrayData as double[])));
            //    Debug.WriteLine("is parallel " + 
            //        (solidUtil.IsCoDirectional(v, v1) || 
            //         solidUtil.IsCoDirectional(solidUtil.Negative(v), v1)));
            //}
        }

        public IFace2 FindRemoteFaceOfCylinderLocator(Locator locator, IFace2 originFace)
        {
            Debug.WriteLine("FindRemoteFaceOfCylinderLocator");
            double[] p = originFace.Normal;
            Debug.WriteLine("origin     " + string.Join(" | ", p));
            IBody2 body = locator.Component.GetBody();
            object[] faces = body.GetFaces();
            IFace2 f = null;

            foreach (object o in faces)
            {
                IFace2 face = o as IFace2;
                //(face as IEntity).Select4(true, null);
                ISurface surface = face.GetSurface();
                if (surface.IsPlane())
                {
                    double[] q = face.Normal;
                    Debug.WriteLine(string.Join(" | ", q));
                    if (p[0] == q[0] && p[1] == q[1] && p[2] == q[2])
                    {
                        return face;
                    }
                    Thread.Sleep(1000);
                }
            }

            return f;
        }

        public void DoProcessSelection()
        {
            ISelectionMgr selector = document.SelectionManager;
            int quantity = selector.GetSelectedObjectCount();
            if (quantity != 3)
            {
                throw new ArgumentException("Make sure, what you pick 3 face!");
            }
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
    }
}
