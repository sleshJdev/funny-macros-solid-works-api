using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace FunnyMacros.Util
{
    class Mounter
    {
        //quantity of errors
        private int qe;

        private Mounter(IModelDoc2 document)
        {
            Document = document;
            Assembly = document as IAssemblyDoc;
        }

        public static Mounter Instance { get; set; }
        private IAssemblyDoc Assembly { get; set; }
        private IModelDoc2 Document { get; set; }

        public static Mounter Initialize(IModelDoc2 document)
        {
            if (Instance == null)
            {
                Instance = new Mounter(document);
            }

            return Instance;
        }

        public void ClearSelection()
        {
            Document.ClearSelection2(true);
        }

        public void AddMate(IFace2 face1, IFace2 face2, int mateType, int alignType)
        {
            ClearSelection();
            (face1 as IEntity).Select4(true, null);
            (face2 as IEntity).Select4(true, null);
            Assembly.AddMate3(mateType, alignType, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            ClearSelection();
        }

        public void AddMate(IFeature feature, IFace2 face, int mateType, int alignType)
        {
            ClearSelection();
            feature.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            (face as IEntity).Select4(true, null);
            Assembly.AddMate3(mateType, alignType, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            ClearSelection();
        }

        public void AddMate(IFeature feature1, IFeature feature2, int mateType, int alignType)
        {
            ClearSelection();
            feature1.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            feature2.Select2(true, (int)swSelectionMarkAction_e.swSelectionMarkAppend);
            Assembly.AddMate3(mateType, alignType, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out qe);
            ClearSelection();
        }
    }
}
