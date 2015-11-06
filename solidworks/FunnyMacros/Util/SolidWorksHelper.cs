using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using FunnyMacros.Model;
using System.Threading;

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
            Debug.WriteLine("loading of assembly ... done!");

            corpus.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CORPUS_NAME));
            planeLocator.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, PLACE_LOCATOR_NAME));
            firstCylinderLocator.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME));
            secondCylinderLocator.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME));
            shaft.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, SHAFT_NAME));
            Debug.WriteLine("loading of models ... done!");

            corpus.Component = LoadUtil.AddModelToAssembly(assembly, corpus.Model.GetPathName(), 0.0, -60.0 / 1000, 0.0);
            planeLocator.Component = LoadUtil.AddModelToAssembly(assembly, planeLocator.Model.GetPathName(), 250.0 / 1000, 0.0, 0.0);
            firstCylinderLocator.Component = LoadUtil.AddModelToAssembly(assembly, firstCylinderLocator.Model.GetPathName(), 0.0, 0.0, 100.0 / 1000);
            secondCylinderLocator.Component = LoadUtil.AddModelToAssembly(assembly, secondCylinderLocator.Model.GetPathName(), 0.0, 0.0, -100.0 / 1000);
            shaft.Component = LoadUtil.AddModelToAssembly(assembly, shaft.Model.GetPathName(), 0.0, 0.0, 0.0);
            Debug.WriteLine("adding to assemply ... done!");

            ClearSelection();
            shaft.Component.Select4(true, null, false);
            assembly.FixComponent();
            Debug.WriteLine("fix shaft ... done!");

            ClearSelection();
            corpus.Component.Select4(true, null, false);
            planeLocator.Component.Select4(true, null, false);
            firstCylinderLocator.Component.Select4(true, null, false);
            secondCylinderLocator.Component.Select4(true, null, false);
            assembly.UnfixComponent();
            Debug.WriteLine("unfix locators  ... done!");

            mathUtility = solidWorks.IGetMathUtility();
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
            AlignWithHorizont(corpus.Component);
            AlignWithHorizont(shaft.Component);
            AlignWithHorizont(planeLocator.Component);
            AlignWithHorizont(firstCylinderLocator.Component);
            AlignWithHorizont(secondCylinderLocator.Component);
        }

        private void AlignWithHorizont(IComponent2 component)
        {
            ClearSelection();
            IFeature face = component.FeatureByName("Сверху");
            IFeature horizont = assembly.IFeatureByName("Сверху");

            face = face == null ? component.FeatureByName("Top") : face;
            horizont = horizont == null ? assembly.IFeatureByName("Top") : horizont;

            face.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            horizont.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);

            int a = (int)swMateType_e.swMatePARALLEL;
            int b = (int)swMateAlign_e.swAlignNONE;
            assembly.AddMate3(a, b, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            ClearSelection();
        }

        public void AddMate()
        {
            IFeature feature = planeLocator.Component.FeatureByName("mouting-plane");
            AddMate(feature, planeBase.Face, (int)swMateType_e.swMateCOINCIDENT);
            Debug.WriteLine("mate for plane base  with shaft ... done!");
            
            AddMateCylinderLocator(firstCylinderLocator.Component, firstCylinderBase.Face);
            AddMateCylinderLocator(secondCylinderLocator.Component, secondCylinderBase.Face);
            Debug.WriteLine("cylinder bases mate with shaft ... done!");

            MiscUtil.Translate(mathUtility, firstCylinderLocator.Component, -0.4, -0.4, -0.4);
            MiscUtil.Translate(mathUtility, secondCylinderLocator.Component, 0.6, 0.6, 0.6);
            Commit();
            Debug.WriteLine("added mate ... done!");
        }

        private void AddMateCylinderLocator(IComponent2 cylinderLocator, IFace2 cylinderBase)
        {
            IFeature firstFeature = cylinderLocator.FeatureByName("mount-plane-1");
            IFeature secondFeature = cylinderLocator.FeatureByName("mount-plane-2");

            AddMate(firstFeature, cylinderBase, (int)swMateType_e.swMateTANGENT);
            AddMate(secondFeature, cylinderBase, (int)swMateType_e.swMateTANGENT);
        }

        private void AddMate(IFeature locator, IFace2 face, int mateType)
        {
            ClearSelection();
            locator.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            (face as IEntity).Select4(true, null);
            assembly.AddMate3(mateType, (int)swMateAlign_e.swAlignSAME, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            ClearSelection();
        }

        public void AlignWithShaft()
        {
            //Bounding(document, MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, planeBase.Component.Transform2, planeBase.Face.GetBox() as double[]));
            //Bounding(document, planeLocator.Component.GetBox(true, true) as double[]);
            double[] planeBaseCenter = MiscUtil.GetCenterOf(planeBase.Face.GetBox() as double[]);
            MiscUtil.Translate(mathUtility, planeLocator.Component,
                planeBaseCenter[0]/*dx*/,
                planeBaseCenter[1]/*dy*/,
                planeBaseCenter[2]/*dz*/);

            double[] firstBox = MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, firstCylinderBase.Component.Transform2, firstCylinderBase.Face.GetBox() as double[]);
            firstBox = MiscUtil.GetCenterOf(firstBox);
            MiscUtil.Translate(mathUtility, firstCylinderLocator.Component, firstBox[0], firstBox[1], firstBox[2]);

            double[] secondBox = MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, secondCylinderBase.Component.Transform2, secondCylinderBase.Face.GetBox() as double[]);
            secondBox = MiscUtil.GetCenterOf(secondBox);
            MiscUtil.Translate(mathUtility, secondCylinderLocator.Component, secondBox[0], secondBox[1], secondBox[2]);

            Commit();
            Debug.WriteLine("align with shaft ... done!");

            //Debug.WriteLine("finding ends of shaft...");
            //RemovalPairlFaces ends = MiscUtil.FindMaximumRemovalPlanes(mathUtility, Array.ConvertAll(shaft.Component.IGetBody().GetFaces(), new Converter<object, IFace2>((o) => { return o as IFace2; })));
            //Debug.WriteLine("ends of shaft search ... done!");

            //double[] firstEndFaceNormal = ends.From.Normal as double[];
            //double[] secondEndFaceNormal = ends.To.Normal as double[];
            //Debug.WriteLine("first base normal: {0}", string.Join(" | ", firstEndFaceNormal), string.Empty);
            //Debug.WriteLine("second base normal: {0}", string.Join(" | ", secondEndFaceNormal), string.Empty);

            //(ends.From as IEntity).Select4(true, null);
            //(ends.To as IEntity).Select4(true, null);
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
                        planeBase = new Base() { Face = face, Component = component };
                    }
                    else if (surface.IsCylinder())
                    {
                        if (firstCylinderBase == null)
                        {
                            firstCylinderBase = new Base() { Face = face, Component = component };
                        }
                        else if(secondCylinderBase == null)
                        {
                            secondCylinderBase = new Base() { Face = face, Component = component };
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
