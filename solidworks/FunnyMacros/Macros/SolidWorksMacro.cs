using System;
using System.Diagnostics;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using FunnyMacros.Model;
using FunnyMacros.Util;
using System.Collections.Generic;

namespace FunnyMacros.Macros
{
    public delegate void ShowInputDialog(string title, string promptText, ref string value);
    public class SolidWorksMacro
    {
        private static string TOP_PLANE_NAME_EN = "Top";
        private static string TOP_PLANE_NAME_RU = "Сверху";

        //quantity of errors
        private static int qe;

        //utils
        private Helper helper;
        private Loader loader;
        private Mounter mounter;

        //propertie: names of parameters of components, references planes
        private IDictionary<Property, string> configuration;

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
        private Locator prismLocator1;
        private Locator prismLocator2;

        public SolidWorksMacro(ISldWorks solidWorks, IDictionary<Property, string> configuration)
        {
            this.solidWorks = solidWorks;
            this.configuration = configuration;
        }

        public bool Run()
        {
            if (Initalize())
            {
                if (DoProcessSelection())
                {
                    try
                    {
                        LoadLocators();
                        AddMate();
                        SetupLocatorsSize();
                        SetupLocatorsLocation();
                        LoadCorpus();
                        SetupCorpus();
                        SetupCorpusSize();
                    }
                    catch (Exception e)
                    {
                        ShowMessage(e.Message);
                        return false;
                    }
                    ResetState();
                    ShowMessage("Well done!");
                    return true;
                };
            };
            return false;
        }

        private bool Initalize()
        {
            helper      = Helper.Initialize(solidWorks.IGetMathUtility());
            corpus      = new Locator(solidWorks.IGetMathUtility());
            shaft       = new Locator(solidWorks.IGetMathUtility());
            planeLocator        = new Locator(solidWorks.IGetMathUtility());
            prismLocator1    = new Locator(solidWorks.IGetMathUtility());
            prismLocator2    = new Locator(solidWorks.IGetMathUtility());

            document = solidWorks.ActiveDoc as IModelDoc2;
            assembly = document as IAssemblyDoc;

            if (assembly == null)
            {
                ShowMessage("Please, open document with assembly");
                return false;
            }

            horizont = assembly.IFeatureByName(TOP_PLANE_NAME_EN);
            horizont = horizont == null ? assembly.IFeatureByName(TOP_PLANE_NAME_RU) : horizont;
            loader = Loader.Initialize(solidWorks, assembly);
            mounter = Mounter.Initialize(document);
            Debug.WriteLine("loading of assembly ... done!");

            object[] components = assembly.GetComponents(true);
            foreach(IComponent2 component in components)
            {
                if (component.Name2.Contains(configuration[Property.MAIN_DETAIL_NAME]))
                {
                    shaft.Component = component;
                    Debug.WriteLine("shaft search ... done!");
                    break;
                }
            }

            if (shaft.Component == null)
            {
                ShowMessage("Please, add shaft to assembly or enter other/correct name");
                return false;
            }

            return true;
        }

        private bool DoProcessSelection()
        {
            ISelectionMgr selectionManager = document.SelectionManager;
            int quantity = selectionManager.GetSelectedObjectCount();
            if (quantity != 3)
            {
                ShowMessage("Make sure, what you pick 3 face and try again");
                return false;
            }
            Debug.WriteLine("selected {0} surface", quantity);

            for (int i = 1; i <= quantity; ++i)
            {
                try
                {
                    IComponent2 component = selectionManager.GetSelectedObjectsComponent4(i, -1);
                    IFace2 face = selectionManager.GetSelectedObject6(i, -1);
                    IFeature feature = face.IGetFeature();
                    ISurface surface = face.IGetSurface();
                    Debug.WriteLine("face {0}. details: materialIdName: {1}, materialUserName: {2}", i, face.MaterialIdName, face.MaterialUserName);
                    Debug.WriteLine("\tfeature details. name: {0}, visible: {1}, description: {2}", feature.Name, feature.Visible, feature.Description);
                    Debug.WriteLine("\tsurface details. isPlane: {0}, isCylinder: {1}", surface.IsPlane(), surface.IsCylinder());
                    MountFace mountFace = new MountFace(solidWorks.IGetMathUtility()) { Face = face, Component = component };
                    if (surface.IsPlane() && planeBase == null)
                    {
                        planeBase = mountFace;
                    }
                    else if (surface.IsCylinder())
                    {
                        if (cylinderBase1 == null)
                        {
                            cylinderBase1 = mountFace;
                        }
                        else if (cylinderBase2 == null)
                        {
                            cylinderBase2 = mountFace;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("process selected faces error: {0}. \n\tstack trace: {1}", ex.Message, ex.StackTrace);
                    return false;
                };
            }
            ClearSelection();
            Debug.WriteLine("processing select of user ... done!");

            return true;
        }

        private void LoadLocators()
        {
            loader.LoadModel(planeLocator.FullPath = GetFile("Choose the plane locator..."));
            planeLocator.Component = loader.AddModelToAssembly(planeLocator.FullPath, 0.0, 0.0, 0.0);
            Activate(document.GetTitle());
            Debug.WriteLine("loading plane locator ... done!");

            loader.LoadModel(prismLocator1.FullPath = GetFile("Choose the first cylinder locator..."));
            prismLocator1.Component = loader.AddModelToAssembly(prismLocator1.FullPath, 0.0, 0.0, 0.0);
            Activate(document.GetTitle());
            Debug.WriteLine("loading cylinder 1 locator ... done!");

            loader.LoadModel(prismLocator2.FullPath = GetFile("Choose the second cylinder locator..."));
            prismLocator2.Component = loader.AddModelToAssembly(prismLocator2.FullPath, 0.0, 0.0, 0.0);
            Activate(document.GetTitle());
            Debug.WriteLine("loading cylinder 2 locator ... done!");

            ClearSelection();
            planeLocator.Component.Select4(true, null, false);
            prismLocator1.Component.Select4(true, null, false);
            prismLocator2.Component.Select4(true, null, false);
            assembly.UnfixComponent();
            Debug.WriteLine("unfix locators  ... done!");

            ClearSelection();
            shaft.Component.Select4(true, null, false);
            assembly.FixComponent();
            Debug.WriteLine("fix shaft ... done!");

            AlignWithPlane(shaft.Component, horizont);
            AlignWithPlane(planeLocator.Component, horizont);
            AlignWithPlane(prismLocator1.Component, horizont);
            AlignWithPlane(prismLocator2.Component, horizont);
            Debug.WriteLine("align with horizont ... done!");
        }

        private void AddMate()
        {
            IFeature feature = planeLocator.Component.FeatureByName(configuration[Property.MOUNTING_PLANE_NAME]);
            mounter.AddMate(feature, planeBase.Face, (int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swAlignSAME);
            Debug.WriteLine("mate for plane base  with shaft ... done!");

            AddMateCylinderLocator(prismLocator1.Component, cylinderBase1.Face);
            AddMateCylinderLocator(prismLocator2.Component, cylinderBase2.Face);
            Debug.WriteLine("cylinder bases mate with shaft ... done!");

            Rebuild();
            Debug.WriteLine("add mate ... done!");
        }

        private void SetupLocatorsSize()
        {
            Box box = planeBase.FaceLBox;
            double width = Math.Max(box.Xmax - box.Xmin, box.Zmax - box.Zmin);
            planeLocator.SetParameter(configuration[Property.PARAMETER_PLANE_LOCATOR_WIDTH], Convert.ToInt32(500.0 * width));
            prismLocator1.SetParameter(configuration[Property.PARAMETER_CYLINDER_LOCATOR_DIAMETER], Convert.ToInt32(2000.0 * cylinderBase1.Radius));
            prismLocator2.SetParameter(configuration[Property.PARAMETER_CYLINDER_LOCATOR_DIAMETER], Convert.ToInt32(2000.0 * cylinderBase2.Radius));

            Rebuild();
            Debug.WriteLine("setup location size ... done!");
        }

        private void SetupLocatorsLocation()
        {
            planeLocator.Translate(planeBase.FaceGBox.Center);
            prismLocator1.Translate(cylinderBase1.FaceGBox.Center);
            prismLocator2.Translate(cylinderBase2.FaceGBox.Center);
            Rebuild();
            Debug.WriteLine("align with shaft ... done!");
        }

        private void LoadCorpus()
        {
            loader.LoadModel(corpus.FullPath = GetFile("Choose the corpus..."));
            corpus.Component = loader.AddModelToAssembly(corpus.FullPath, 0.0, 0.0, 0.0);
            Debug.WriteLine("loading of corpus ... done!");

            AlignWithPlane(corpus.Component, horizont);
            Debug.WriteLine("align of corpus ... done!");

            ClearSelection();
            corpus.Component.Select4(true, null, false);
            assembly.UnfixComponent();
            Debug.WriteLine("unfix corpus  ... done!");

            Activate(document.GetTitle());
        }

        private void SetupCorpus()
        {
            IFace2 face1 = prismLocator1.BottomFace;
            IFace2 face2 = prismLocator2.BottomFace;
            (face1 as Entity).Select4(true, null);
            (face2 as Entity).Select4(true, null);
            IFace2 lowestFace = null;
            Locator danglingLocator = null;
            if (helper.CenterBox(face1.GetBox()).Y < helper.CenterBox(face2.GetBox()).Y)
            {
                lowestFace = face1;
                danglingLocator = prismLocator2;
            }
            else
            {
                lowestFace = face2;
                danglingLocator = prismLocator1;
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
            value = delta + Convert.ToDouble(danglingLocator.GetParameter(configuration[Property.PARAMETER_CYLINDER_LOCATOR_HEIGHT]));
            danglingLocator.SetParameter(configuration[Property.PARAMETER_CYLINDER_LOCATOR_HEIGHT], Convert.ToInt32(value).ToString());
            Debug.WriteLine("extension of the dangling locator, delta: {0}mm, value: {1} ... done!", delta, Convert.ToInt32(value));

            delta = Math.Abs(1000.0 * (planeBase.FaceGBox.Ymax - planeLocator.LBox.Ymax));
            value = Convert.ToDouble(planeLocator.GetParameter(configuration[Property.PARAMETER_PLANE_LOCATOR_HEIGHT])) + delta;
            planeLocator.SetParameter(configuration[Property.PARAMETER_PLANE_LOCATOR_HEIGHT], Convert.ToInt32(value).ToString());
            Debug.WriteLine("extension of the plane locator, delta {0}mm, value: {1}mm ... done!", delta, Convert.ToInt32(value));

            Rebuild();

            Debug.WriteLine("setup corpus ... done!");
        }

        public void SetupCorpusSize()
        {
            int i;
            Box[] boxes = { planeLocator.LBox, prismLocator1.LBox, prismLocator2.LBox };
            Box boundingBox = new Box();
            i = 0; boundingBox.Xmin = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i] }.Min();
            i = 1; boundingBox.Ymin = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i] }.Min();
            i = 2; boundingBox.Zmin = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i] }.Min();
            i = 3; boundingBox.Xmax = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i] }.Max();
            i = 4; boundingBox.Ymax = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i] }.Max();
            i = 5; boundingBox.Zmax = new double[] { boxes[0].CoordinatesCorners[i], boxes[1].CoordinatesCorners[i], boxes[2].CoordinatesCorners[i] }.Max();

            int width = Convert.ToInt32(1000.0 * (boundingBox.Zmax - boundingBox.Zmin));
            int length = Convert.ToInt32(1000.0 * (boundingBox.Xmax - boundingBox.Xmin));
            int size = Math.Max(width, length);

            Vector newLocation = new Vector()
            {
                X = (boundingBox.Xmin + boundingBox.Xmax) / 2,
                Y = (boundingBox.Ymin + boundingBox.Ymax) / 2,
                Z = (boundingBox.Zmin + boundingBox.Zmax) / 2
            };

            corpus.Translate(newLocation);
            corpus.SetParameter(configuration[Property.PARAMETER_CORPUS_WIDTH], size.ToString());
            corpus.SetParameter(configuration[Property.PARAMETER_CORPUS_LENGTH], size.ToString());

            Rebuild();
        }

        private void AlignWithPlane(IComponent2 component, IFeature baseFace)
        {
            IFeature face = component.FeatureByName(TOP_PLANE_NAME_RU);
            face = face == null ? component.FeatureByName(TOP_PLANE_NAME_EN) : face;
            mounter.AddMate(baseFace, face, (int)swMateType_e.swMatePARALLEL, (int)swMateAlign_e.swAlignNONE);
        }

        private void AddMateCylinderLocator(IComponent2 cylinderLocator, IFace2 cylinderBase)
        {
            IFeature firstFeature = cylinderLocator.FeatureByName(configuration[Property.MOUNTING_CYLINDER_PLANE_NAME_1]);
            IFeature secondFeature = cylinderLocator.FeatureByName(configuration[Property.MOUNTING_CYLINDER_PLANE_NAME_2]);

            mounter.AddMate(firstFeature, cylinderBase, (int)swMateType_e.swMateTANGENT, (int)swMateAlign_e.swAlignSAME);
            mounter.AddMate(secondFeature, cylinderBase, (int)swMateType_e.swMateTANGENT, (int)swMateAlign_e.swAlignSAME);
        }

        private string GetFile(string title)
        {
            string filters = "SolidWorks Files (*.sldprt; *.sldasm; *.slddrw)|*.sldprt;*.sldasm;*.slddrw|Filter name (*.fil)|*.fil|All Files (*.*)|*.*|";
            int fileOpenOptions;
            string fileConfigName;
            string fileDisplayName;
            solidWorks.GetOpenFileName(title, string.Empty, filters, out fileOpenOptions, out fileConfigName, out fileDisplayName);

            return fileDisplayName;
        }

        private void ResetState()
        {
            planeBase = cylinderBase1 = cylinderBase2 = null;
            planeLocator = prismLocator1 = prismLocator2 = null;
            corpus = null;
            shaft = null;
            horizont = null;
        }

        private void ShowMessage(string message)
        {
            solidWorks.SendMsgToUser2(message, (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOk);
        }

        private void Activate(string name)
        {
            solidWorks.ActivateDoc3(document.GetTitle(), false, (int)swRebuildOnActivation_e.swUserDecision, ref qe);
        }

        private void ClearSelection()
        {
            document.ClearSelection2(true);
        }

        private void Rebuild()
        {
            document.EditRebuild3();
        }
    }
}
