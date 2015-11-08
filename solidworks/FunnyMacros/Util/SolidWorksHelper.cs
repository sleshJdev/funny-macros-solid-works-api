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

        //utils
        private Helper helper;
        private Mounter mounter;

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

            corpus  = new Locator(mathutility);
            shaft   = new Locator(mathutility);

            planeLocator        = new Locator(mathutility);
            cylinderLocator1    = new Locator(mathutility);
            cylinderLocator2    = new Locator(mathutility);

            Loader loader = Loader.Initialize();
            loader.SolidWorks = solidWorks;
            document = loader.LoadAssembly(Path.Combine(Settings.HOME_PATH, Settings.ASSEMBLY_NAME));
            assembly = document as IAssemblyDoc;
            loader.Assembly = assembly;
            mounter  = Mounter.Initialize(document);
            horizont = assembly.IFeatureByName(Settings.TOP_PLANE_NAME_RU);
            horizont = horizont == null ? assembly.IFeatureByName(Settings.TOP_PLANE_NAME_EN) : horizont;
            Debug.WriteLine("loading of assembly ... done!");

            corpus.Model            = loader.LoadModel(Path.Combine(Settings.HOME_PATH, Settings.CORPUS_NAME));
            planeLocator.Model      = loader.LoadModel(Path.Combine(Settings.HOME_PATH, Settings.PLACE_LOCATOR_NAME));
            cylinderLocator1.Model  = loader.LoadModel(Path.Combine(Settings.HOME_PATH, Settings.CYLINDER_LOCATOR_NAME_1));
            cylinderLocator2.Model  = loader.LoadModel(Path.Combine(Settings.HOME_PATH, Settings.CYLINDER_LOCATOR_NAME_2));
            shaft.Model             = loader.LoadModel(Path.Combine(Settings.HOME_PATH, Settings.SHAFT_NAME));
            Debug.WriteLine("loading of models ... done!");

            corpus.Component            = loader.AddModelToAssembly(corpus.Model.GetPathName(), 0.0, -60.0 / 1000, 0.0);
            planeLocator.Component      = loader.AddModelToAssembly(planeLocator.Model.GetPathName(), 250.0 / 1000, 0.0, 0.0);
            cylinderLocator1.Component  = loader.AddModelToAssembly(cylinderLocator1.Model.GetPathName(), 0.0, 0.0, 150.0 / 1000);
            cylinderLocator2.Component  = loader.AddModelToAssembly(cylinderLocator2.Model.GetPathName(), 0.0, 0.0, -150.0 / 1000);
            shaft.Component             = loader.AddModelToAssembly(shaft.Model.GetPathName(), 0.0, 0.0, 0.0);
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

        private void Rebuild()
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

        private void AlignWithPlane(IComponent2 component, IFeature baseFace)
        {
            IFeature face = component.FeatureByName(Settings.TOP_PLANE_NAME_RU);
            face = face == null ? component.FeatureByName(Settings.TOP_PLANE_NAME_EN) : face;
            mounter.AddMate(baseFace, face, (int)swMateType_e.swMatePARALLEL, (int)swMateAlign_e.swAlignNONE);
        }

        public void AddMate()
        {
            IFeature feature = planeLocator.Component.FeatureByName(Settings.MOUNTING_PLANE_NAME);
            mounter.AddMate(feature, planeBase.Face, (int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swAlignSAME);
            Debug.WriteLine("mate for plane base  with shaft ... done!");
            
            AddMateCylinderLocator(cylinderLocator1.Component, cylinderBase1.Face);
            AddMateCylinderLocator(cylinderLocator2.Component, cylinderBase2.Face);
            Debug.WriteLine("cylinder bases mate with shaft ... done!");

            cylinderLocator1.Translate(new Vector(-0.4, -0.4, -0.4));
            cylinderLocator2.Translate(new Vector(0.6, 0.6, 0.6));

            Rebuild();
            Debug.WriteLine("added mate ... done!");
        }

        private void AddMateCylinderLocator(IComponent2 cylinderLocator, IFace2 cylinderBase)
        {
            IFeature firstFeature = cylinderLocator.FeatureByName(Settings.MOUNTING_CYLINDER_PLANE_NAME_1);
            IFeature secondFeature = cylinderLocator.FeatureByName(Settings.MOUNTING_CYLINDER_PLANE_NAME_2);

            mounter.AddMate(firstFeature, cylinderBase, (int)swMateType_e.swMateTANGENT, (int)swMateAlign_e.swAlignSAME);
            mounter.AddMate(secondFeature, cylinderBase, (int)swMateType_e.swMateTANGENT, (int)swMateAlign_e.swAlignSAME);
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
            mounter.AddMate(topFaceOfCorpus, lowestFace, (int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swAlignAGAINST);
            mounter.AddMate(topFaceOfCorpus, planeLocator.BottomFace, (int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swAlignAGAINST);
            Debug.WriteLine("mate with corpus of not dangling locator ... done!");

            double delta = 0;
            double value = 0;

            Box boxLowestFaceDanglingLocator = new Box(helper.ApplyTransform(danglingLocator.Transform, danglingLocator.BottomFace.GetBox()));
            Box boxTopFaceCorpus = new Box(helper.ApplyTransform(corpus.Transform, topFaceOfCorpus.GetBox()));

            delta = Math.Abs(1000.0 * (boxLowestFaceDanglingLocator.Ymin - boxTopFaceCorpus.Ymin));
            value = delta + Convert.ToDouble(helper.GetParameter(danglingLocator.EquationManager, Settings.PARAMETER_CYLINDER_LOCATOR_HEIGHT));
            helper.SetParameter(danglingLocator.EquationManager, Settings.PARAMETER_CYLINDER_LOCATOR_HEIGHT, Convert.ToInt32(value).ToString());
            Debug.WriteLine("extension of the dangling locator, delta: {0}mm, value: {1} ... done!", delta, Convert.ToInt32(value));

            delta = Math.Abs(1000.0 * (planeBase.FaceGBox.Ymax - planeLocator.LBox.Ymax));
            value = Convert.ToDouble(helper.GetParameter(planeLocator.EquationManager, Settings.PARAMETER_PLANE_LOCATOR_HEIGHT)) + delta;
            helper.SetParameter(planeLocator.EquationManager, Settings.PARAMETER_PLANE_LOCATOR_HEIGHT, Convert.ToInt32(value).ToString());
            Debug.WriteLine("extension of the plane locator, delta {0}mm, value: {1}mm ... done!", delta, Convert.ToInt32(value));

            Rebuild();

            Debug.WriteLine("setup corpus ... done!");
        }

        public void AdjustSize()
        {
            int i;
            Box[] boxes = { shaft.LBox, planeLocator.LBox, cylinderLocator1.LBox, cylinderLocator2.LBox };
            Box boundingBox = new Box();
            i = 0; boundingBox.Xmin = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i], boxes[3].CoordinatesCorners[i] }.Min();
            i = 1; boundingBox.Ymin = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i], boxes[3].CoordinatesCorners[i] }.Min();
            i = 2; boundingBox.Zmin = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i], boxes[3].CoordinatesCorners[i] }.Min();
            i = 3; boundingBox.Xmax = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i], boxes[3].CoordinatesCorners[i] }.Max();
            i = 4; boundingBox.Ymax = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i], boxes[3].CoordinatesCorners[i] }.Max();
            i = 5; boundingBox.Zmax = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i], boxes[3].CoordinatesCorners[i] }.Max();

            int width = Convert.ToInt32(1000.0 * (boundingBox.Zmax - boundingBox.Zmin)) + 50;
            int length = Convert.ToInt32(1000.0 * (boundingBox.Xmax - boundingBox.Xmin)) + 50;

            helper.SetParameter(corpus.EquationManager, Settings.PARAMETER_CORPUS_WIDTH, width.ToString());
            helper.SetParameter(corpus.EquationManager, Settings.PARAMETER_CORPUS_LENGTH, length.ToString());

            Rebuild();
        }

        public void AlignWithShaft()
        {
            planeLocator.Translate(planeBase.CenterFaceGBox);
            cylinderLocator1.Translate(cylinderBase1.CenterFaceGBox);
            cylinderLocator2.Translate(cylinderBase2.CenterFaceGBox);
            Rebuild();
            Debug.WriteLine("align with shaft ... done!");
        }

        public void DoProcessSelection()
        {
            ISelectionMgr selectionManager = document.SelectionManager;
            int quantity = selectionManager.GetSelectedObjectCount();
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
                    IComponent2 component = selectionManager.GetSelectedObjectsComponent4(i, -1) as IComponent2;
                    IFace2 face = selectionManager.GetSelectedObject6(i, -1) as IFace2;
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
