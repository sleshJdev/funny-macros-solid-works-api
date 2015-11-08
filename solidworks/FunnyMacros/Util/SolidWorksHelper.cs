using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using FunnyMacros.Model;
using System.Threading;
using System.Text.RegularExpressions;
using System.Linq;

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
        private const string CYLINDER_LOCATOR_NAME_1 = "cylinder-locator-1.sldprt";
        private const string CYLINDER_LOCATOR_NAME_2 = "cylinder-locator-2.sldprt";

        private const string PARAMETER_CORPUS_WIDTH = "B";
        private const string PARAMETER_CORPUS_LENGTH = "L";
        private const string PARAMETER_PLANE_LOCATOR_HEIGHT = "L";
        private const string PARAMETER_CYLINDER_LOCATOR_HEIGHT = "___height_cilynder_locator_base___";

        //utils
        private Helper helper;

        //solidworks application instance
        public ISldWorks solidWorks;

        //the assembly of details
        private IModelDoc2 document;
        private IAssemblyDoc assembly;
        private IFeature horizont;

        //the corpus - the base for the design
        private Locator corpus;

        //the detail for positionign
        private Locator shaft;

        //bases of that detail
        private MountFace planeBase;
        private MountFace cylinderBase1;
        private MountFace cylinderBase2;

        //locator for bases
        private Locator planeLocator;
        private Locator cylinderLocator1;
        private Locator cylinderLocator2;

        public void Open()
        {
            solidWorks = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("SldWorks.Application")) as ISldWorks;
            solidWorks.Visible = true;
        }

        public void Initalize()
        {
            IMathUtility mathutility = solidWorks.IGetMathUtility();
            helper = Helper.Initialize(mathutility);

            corpus = new Locator(mathutility);
            shaft = new Locator(mathutility);

            planeLocator = new Locator(mathutility);
            cylinderLocator1 = new Locator(mathutility);
            cylinderLocator2 = new Locator(mathutility);

            document = Loader.LoadAssembly(solidWorks, Path.Combine(HOME_PATH, ASSEMBLY_NAME));
            assembly = document as IAssemblyDoc;
            horizont = assembly.IFeatureByName("Сверху");
            horizont = horizont == null ? assembly.IFeatureByName("Top") : horizont;
            Debug.WriteLine("loading of assembly ... done!");

            corpus.Model = Loader.LoadModel(solidWorks, Path.Combine(HOME_PATH, CORPUS_NAME));
            planeLocator.Model = Loader.LoadModel(solidWorks, Path.Combine(HOME_PATH, PLACE_LOCATOR_NAME));
            cylinderLocator1.Model = Loader.LoadModel(solidWorks, Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME_1));
            cylinderLocator2.Model = Loader.LoadModel(solidWorks, Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME_2));
            shaft.Model = Loader.LoadModel(solidWorks, Path.Combine(HOME_PATH, SHAFT_NAME));
            Debug.WriteLine("loading of models ... done!");

            corpus.Component = Loader.AddModelToAssembly(assembly, corpus.Model.GetPathName(), 0.0, -60.0 / 1000, 0.0);
            planeLocator.Component = Loader.AddModelToAssembly(assembly, planeLocator.Model.GetPathName(), 250.0 / 1000, 0.0, 0.0);
            cylinderLocator1.Component = Loader.AddModelToAssembly(assembly, cylinderLocator1.Model.GetPathName(), 0.0, 0.0, 150.0 / 1000);
            cylinderLocator2.Component = Loader.AddModelToAssembly(assembly, cylinderLocator2.Model.GetPathName(), 0.0, 0.0, -150.0 / 1000);
            shaft.Component = Loader.AddModelToAssembly(assembly, shaft.Model.GetPathName(), 0.0, 0.0, 0.0);
            Debug.WriteLine("adding to assemply ... done!");

            ClearSelection();
            shaft.Component.Select4(true, null, false);
            assembly.FixComponent();
            Debug.WriteLine("fix shaft ... done!");

            ClearSelection();
            corpus.Component.Select4(true, null, false);
            planeLocator.Component.Select4(true, null, false);
            cylinderLocator1.Component.Select4(true, null, false);
            cylinderLocator2.Component.Select4(true, null, false);
            assembly.UnfixComponent();
            Debug.WriteLine("unfix locators  ... done!");
        }

        private void ClearSelection()
        {
            document.ClearSelection2(true);
        }

        private void Commit()
        {
            document.EditRebuild3();
        }

        public void AlignAllWithHorizont()
        {
            //conjugation with the horizontal plane
            AlignWithPlane(corpus.Component, horizont);
            AlignWithPlane(shaft.Component, horizont);
            AlignWithPlane(planeLocator.Component, horizont);
            AlignWithPlane(cylinderLocator1.Component, horizont);
            AlignWithPlane(cylinderLocator2.Component, horizont);
        }

        private void AlignWithPlane(IComponent2 component, IFeature plane)
        {
            ClearSelection();
            IFeature face = component.FeatureByName("Сверху");
            face = face == null ? component.FeatureByName("Top") : face;

            face.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            plane.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);

            int a = (int)swMateType_e.swMatePARALLEL;
            int b = (int)swMateAlign_e.swAlignNONE;
            assembly.AddMate3(a, b, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            ClearSelection();
        }

        public void AddMate()
        {
            IFeature feature = planeLocator.Component.FeatureByName("___mouting_plane___");
            AddMate(feature, planeBase.Face, (int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swAlignSAME);
            Debug.WriteLine("mate for plane base  with shaft ... done!");
            
            AddMateCylinderLocator(cylinderLocator1.Component, cylinderBase1.Face);
            AddMateCylinderLocator(cylinderLocator2.Component, cylinderBase2.Face);
            Debug.WriteLine("cylinder bases mate with shaft ... done!");

            cylinderLocator1.Translate(new Vector(-0.4, -0.4, -0.4));
            cylinderLocator2.Translate(new Vector(0.6, 0.6, 0.6));

            Commit();
            Debug.WriteLine("added mate ... done!");
        }

        private void AddMateCylinderLocator(IComponent2 cylinderLocator, IFace2 cylinderBase)
        {
            IFeature firstFeature = cylinderLocator.FeatureByName("___mount_plane_1___");
            IFeature secondFeature = cylinderLocator.FeatureByName("___mount_plane_2___");

            AddMate(firstFeature, cylinderBase, (int)swMateType_e.swMateTANGENT, (int)swMateAlign_e.swAlignSAME);
            AddMate(secondFeature, cylinderBase, (int)swMateType_e.swMateTANGENT, (int)swMateAlign_e.swAlignSAME);
        }

        private void AddMate(IFace2 face1, IFace2 face2, int mateType, int alignType)
        {
            ClearSelection();
            (face1 as IEntity).Select4(true, null);
            (face2 as IEntity).Select4(true, null);
            assembly.AddMate3(mateType, alignType, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            ClearSelection();
        }

        private void AddMate(IFeature feature, IFace2 face, int mateType, int alignType)
        {
            ClearSelection();
            feature.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            (face as IEntity).Select4(true, null);
            assembly.AddMate3(mateType, alignType, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            ClearSelection();
        }

        public void SetupCorpus()
        {
            IFace2 lowestFace = null;
            Locator danglingLocator = null;
            if (cylinderLocator1.CenterGBox.Y < cylinderLocator2.CenterGBox.Y)
            {
                lowestFace = cylinderLocator1.BottomFace;
                danglingLocator = cylinderLocator2;
            }
            else
            {
                lowestFace = cylinderLocator2.BottomFace;
                danglingLocator = cylinderLocator1;
            }
            Debug.WriteLine("dangling locator search ... done!");

            IFace2 topFaceOfCorpus = corpus.TopFace;
            AddMate(topFaceOfCorpus, lowestFace, (int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swAlignAGAINST);
            AddMate(topFaceOfCorpus, planeLocator.BottomFace, (int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swAlignAGAINST);
            Debug.WriteLine("mate with corpus of not dangling locator ... done!");

            double delta = 0;
            double value = 0;

            Box boxLowestFaceDanglingLocator = new Box(helper.ApplyTransform(danglingLocator.Transform, danglingLocator.BottomFace.GetBox()));
            Box boxTopFaceCorpus = new Box(helper.ApplyTransform(corpus.Transform, topFaceOfCorpus.GetBox()));
            delta = Math.Abs(1000.0 * (boxLowestFaceDanglingLocator.Ymin - boxTopFaceCorpus.Ymin));
            value = delta + Convert.ToDouble(helper.GetPropertyValue(danglingLocator.EquationManager, PARAMETER_CYLINDER_LOCATOR_HEIGHT));
            helper.SetPropertyValue(danglingLocator.EquationManager, PARAMETER_CYLINDER_LOCATOR_HEIGHT, Convert.ToInt32(value).ToString());
            Debug.WriteLine("extension of the dangling locator, delta: {0}mm, value: {1} ... done!", delta, Convert.ToInt32(value));

            delta = Math.Abs(1000.0 * (planeBase.FaceGBox.Ymax - planeLocator.LBox.Ymax));
            value = Convert.ToDouble(helper.GetPropertyValue(planeLocator.EquationManager, PARAMETER_PLANE_LOCATOR_HEIGHT)) + delta;
            helper.SetPropertyValue(planeLocator.EquationManager, PARAMETER_PLANE_LOCATOR_HEIGHT, Convert.ToInt32(value).ToString());
            Debug.WriteLine("extension of the plane locator, delta {0}mm, value: {1}mm ... done!", delta, Convert.ToInt32(value));

            Commit();

            Debug.WriteLine("setup corpus ... done!");
        }

        public void AdjustSize()
        {
            int i;
            Box[] boxs = { shaft.LBox, planeLocator.LBox, cylinderLocator1.LBox, cylinderLocator2.LBox };
            Box boundingBox = new Box();
            i = 0; boundingBox.Xmin = new double[] { boxs[0].CoordinatesCorners[i], boxs[1].CoordinatesCorners[i], boxs[2].CoordinatesCorners[i], boxs[3].CoordinatesCorners[i] }.Min();
            i = 1; boundingBox.Ymin = new double[] { boxs[0].CoordinatesCorners[i], boxs[1].CoordinatesCorners[i], boxs[2].CoordinatesCorners[i], boxs[3].CoordinatesCorners[i] }.Min();
            i = 2; boundingBox.Zmin = new double[] { boxs[0].CoordinatesCorners[i], boxs[1].CoordinatesCorners[i], boxs[2].CoordinatesCorners[i], boxs[3].CoordinatesCorners[i] }.Min();
            i = 3; boundingBox.Xmax = new double[] { boxs[0].CoordinatesCorners[i], boxs[1].CoordinatesCorners[i], boxs[2].CoordinatesCorners[i], boxs[3].CoordinatesCorners[i] }.Max();
            i = 4; boundingBox.Ymax = new double[] { boxs[0].CoordinatesCorners[i], boxs[1].CoordinatesCorners[i], boxs[2].CoordinatesCorners[i], boxs[3].CoordinatesCorners[i] }.Max();
            i = 5; boundingBox.Zmax = new double[] { boxs[0].CoordinatesCorners[i], boxs[1].CoordinatesCorners[i], boxs[2].CoordinatesCorners[i], boxs[3].CoordinatesCorners[i] }.Max();

            int width = Convert.ToInt32(1000.0 * (boundingBox.Zmax - boundingBox.Zmin)) + 50;
            int length = Convert.ToInt32(1000.0 * (boundingBox.Xmax - boundingBox.Xmin)) + 50;

            helper.SetPropertyValue(corpus.EquationManager, PARAMETER_CORPUS_WIDTH, width.ToString());
            helper.SetPropertyValue(corpus.EquationManager, PARAMETER_CORPUS_LENGTH, length.ToString());

            Commit();
        }

        public void AlignWithShaft()
        {
            planeLocator.Translate(planeBase.CenterFaceGBox);
            cylinderLocator1.Translate(cylinderBase1.CenterFaceGBox);
            cylinderLocator2.Translate(cylinderBase2.CenterFaceGBox);
            Commit();
            Debug.WriteLine("align with shaft ... done!");
        }

        public void DoProcessSelection()
        {
            ISelectionMgr selector = document.SelectionManager;
            int quantity = selector.GetSelectedObjectCount();
            if (quantity != 3)
            {
                solidWorks.SendMsgToUser2("Make sure, what you pick 3 face!", (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOk);
                return;
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
                        planeBase = new MountFace(solidWorks.IGetMathUtility()) { Face = face, Component = component };
                    }
                    else if (surface.IsCylinder())
                    {
                        if (cylinderBase1 == null)
                        {
                            cylinderBase1 = new MountFace(solidWorks.IGetMathUtility()) { Face = face, Component = component };
                        }
                        else if(cylinderBase2 == null)
                        {
                            cylinderBase2 = new MountFace(solidWorks.IGetMathUtility()) { Face = face, Component = component };
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("process selected faces error: {0}. \n\tstack trace: {1}", ex.Message, ex.StackTrace);
                };
            }
            ClearSelection();
            Debug.WriteLine("processing select of user ... done!");
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
