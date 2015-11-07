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

        private const string PARAMETER_PLANE_LOCATOR_HEIGHT = "L";
        private const string PARAMETER_CYLINDER_LOCATOR_HEIGHT = "___height_cilynder_locator_base___";

        //solidworks application instance
        public ISldWorks solidWorks;
        public IMathUtility mathUtility;

        //the assembly of details
        private IModelDoc2 document;
        private IAssemblyDoc assembly;
        private IFeature horizont;

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
        private Locator cylinderLocator1 = new Locator();
        private Locator cylinderLocator2 = new Locator();

        public void Open()
        {
            solidWorks = System.Activator.CreateInstance(System.Type.GetTypeFromProgID("SldWorks.Application")) as ISldWorks;
            solidWorks.Visible = true;
        }

        public void Initalize()
        {
            mathUtility = solidWorks.IGetMathUtility();

            document = LoadUtil.LoadAssembly(solidWorks, Path.Combine(HOME_PATH, ASSEMBLY_NAME));
            assembly = document as IAssemblyDoc;
            horizont = assembly.IFeatureByName("Сверху");
            horizont = horizont == null ? assembly.IFeatureByName("Top") : horizont;
            Debug.WriteLine("loading of assembly ... done!");

            corpus.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CORPUS_NAME));
            planeLocator.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, PLACE_LOCATOR_NAME));
            cylinderLocator1.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME_1));
            cylinderLocator2.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, CYLINDER_LOCATOR_NAME_2));
            shaft.Model = LoadUtil.LoadModel(solidWorks, Path.Combine(HOME_PATH, SHAFT_NAME));
            Debug.WriteLine("loading of models ... done!");

            corpus.Component = LoadUtil.AddModelToAssembly(assembly, corpus.Model.GetPathName(), 0.0, -60.0 / 1000, 0.0);
            planeLocator.Component = LoadUtil.AddModelToAssembly(assembly, planeLocator.Model.GetPathName(), 250.0 / 1000, 0.0, 0.0);
            cylinderLocator1.Component = LoadUtil.AddModelToAssembly(assembly, cylinderLocator1.Model.GetPathName(), 0.0, 0.0, 150.0 / 1000);
            cylinderLocator2.Component = LoadUtil.AddModelToAssembly(assembly, cylinderLocator2.Model.GetPathName(), 0.0, 0.0, -150.0 / 1000);
            shaft.Component = LoadUtil.AddModelToAssembly(assembly, shaft.Model.GetPathName(), 0.0, 0.0, 0.0);
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
            
            AddMateCylinderLocator(cylinderLocator1.Component, firstCylinderBase.Face);
            AddMateCylinderLocator(cylinderLocator2.Component, secondCylinderBase.Face);
            Debug.WriteLine("cylinder bases mate with shaft ... done!");

            MiscUtil.Translate(mathUtility, cylinderLocator1.Component, -0.4, -0.4, -0.4);
            MiscUtil.Translate(mathUtility, cylinderLocator2.Component, 0.6, 0.6, 0.6);
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
            IFace2 face1 = cylinderLocator1.GetBottom(mathUtility);
            IFace2 face2 = cylinderLocator2.GetBottom(mathUtility);

            double[] centerOfFace1 = MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, cylinderLocator1.Transform, MiscUtil.GetCenterOf(face1.GetBox()));
            double[] centerOfFace2 = MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, cylinderLocator2.Transform, MiscUtil.GetCenterOf(face2.GetBox()));

            Debug.WriteLine("center of face2: {0}", string.Join(" | ", centerOfFace1), string.Empty);
            Debug.WriteLine("center of face2: {0}", string.Join(" | ", centerOfFace2), string.Empty);

            IFace2 lowestFaceOfCilynderLocator = null;
            Locator shortLocator = null;
            if (centerOfFace1[1] < centerOfFace2[1])
            {
                lowestFaceOfCilynderLocator = face1;
                shortLocator = cylinderLocator2;
            }
            else
            {
                lowestFaceOfCilynderLocator = face2;
                shortLocator = cylinderLocator1;
            }

            IFace2 topCorpusFace = corpus.GetTop(mathUtility);
            IFace2 lowestFaceOfPlaneLocatorFace = planeLocator.GetBottom(mathUtility);

            int a = (int)swMateType_e.swMateCOINCIDENT;
            int b = (int)swMateAlign_e.swAlignAGAINST;
            AddMate(topCorpusFace, lowestFaceOfCilynderLocator, a, b);
            AddMate(topCorpusFace, lowestFaceOfPlaneLocatorFace, a, b);

            double[] centerLowestFaceShortLocator = MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, shortLocator.Transform, shortLocator.GetBottom(mathUtility).GetBox() as double[]);
            double[] centerLowestFaceCorpus = MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, corpus.Transform, topCorpusFace.GetBox() as double[]);
            double delta = Math.Abs(1000.0 * (centerLowestFaceShortLocator[1] - centerLowestFaceCorpus[1]));

            IEquationMgr equationManager = shortLocator.Model.GetEquationMgr();
            double value = Convert.ToDouble(GetPropertyValue(equationManager, PARAMETER_CYLINDER_LOCATOR_HEIGHT)) + delta;
             
            SetPropertyValue(equationManager, PARAMETER_CYLINDER_LOCATOR_HEIGHT, Convert.ToInt32(value).ToString());
            Commit();

            //AddMate(topCorpusFace, shortLocator.GetBottom(mathUtility), a, b);

            delta = Math.Abs(1000.0 * (MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, planeBase.Transform, planeBase.Box)[4] - 
                                       MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, planeLocator.Transform, planeLocator.Box)[4]));
            value = Convert.ToDouble(GetPropertyValue(equationManager, PARAMETER_PLANE_LOCATOR_HEIGHT)) + delta;
            equationManager = planeLocator.Model.GetEquationMgr();
            SetPropertyValue(equationManager, PARAMETER_PLANE_LOCATOR_HEIGHT, Convert.ToInt32(value).ToString());
            Commit();
        }
 
        public void SetPropertyValue(IEquationMgr manager, string property, string value)
        {
            Debug.WriteLine("file path: {0}", manager.FilePath, string.Empty);
            for (int i = 0; i < manager.GetCount(); ++i)
            {
                string equation = manager.Equation[i];
                Debug.WriteLine("equation {0}, statement: {1}, value: {2}", i, equation, manager.Value[i]);
                if (equation.Contains(property))
                {
                    string newEquation = new Regex(@"=\s*(\d*)").Replace(equation, (m) => { return string.Format("={0}", value); }, 1);
                    Debug.WriteLine("equation: {0}, new equation: {1}", equation, newEquation);
                    manager.Delete(i);
                    manager.Add2(i, newEquation, true);
                    return;
                }
            }
        }

        public string GetPropertyValue(IEquationMgr manager, string property)
        {
            for (int i = 0; i < manager.GetCount(); ++i)
            {
                string equation = manager.Equation[i];
                if (equation.Contains(property))
                {
                    Match match = new Regex(@"=\s*(\d*)").Match(equation);
                    if (match.Success)
                    {
                        return match.Groups[1].Value;
                    }
                }
            }

            return null;
        }

        public void AlignWithShaft()
        {
            //Bounding(document, MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, planeBase.Component.Transform2, planeBase.Face.GetBox() as double[]));
            //Bounding(document, planeLocator.Component.GetBox(true, true) as double[]);
            double[] planeBaseCenter = MiscUtil.GetCenterOf(planeBase.Box);
            MiscUtil.Translate(mathUtility, planeLocator.Component,
                planeBaseCenter[0]/*dx*/,
                planeBaseCenter[1]/*dy*/,
                planeBaseCenter[2]/*dz*/);

            double[] firstBox = MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, firstCylinderBase.Transform, firstCylinderBase.Box);
            firstBox = MiscUtil.GetCenterOf(firstBox);
            MiscUtil.Translate(mathUtility, cylinderLocator1.Component, firstBox[0], firstBox[1], firstBox[2]);

            double[] secondBox = MiscUtil.GetInGlobalCoordinatesSystem(mathUtility, secondCylinderBase.Transform, secondCylinderBase.Box);
            secondBox = MiscUtil.GetCenterOf(secondBox);
            MiscUtil.Translate(mathUtility, cylinderLocator2.Component, secondBox[0], secondBox[1], secondBox[2]);

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
