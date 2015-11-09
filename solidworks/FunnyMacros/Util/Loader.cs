using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace FunnyMacros.Util
{
    class Loader
    {
        //quantity of errors
        private static int qe;

        //quantity of warnings
        private static int qw;

        public static Loader Instance { get; set; }
        public ISldWorks SolidWorks { get; set; }
        public IAssemblyDoc Assembly { get; set; }

        public static Loader Initialize()
        {
            if (Instance == null)
            {
                Instance = new Loader();
            }

            return Instance;
        }

        public IModelDoc2 LoadAssembly(string path)
        {
            int a = (int)swDocTemplateTypes_e.swDocTemplateTypePART;
            IModelDoc2 model = SolidWorks.OpenDoc6(path, a, 0, string.Empty, ref qe, ref qw);
            Debug.WriteLine("assemly loaded. errors: {2}, warnings: {3}, pathName: {0}, title: {1} ... ok", model.GetPathName(), model.GetTitle(), qe, qw);

            SolidWorks.ActivateDoc2(model.GetTitle(), false, ref qe);
            Debug.WriteLine("activate doc. errors: {0} ... ok", qe, string.Empty);

            model = SolidWorks.ActiveDoc as IModelDoc2;
            Debug.WriteLine("set assembly to active document ... ok");

            return model;
        }

        public IModelDoc2 LoadModel(string path)
        {
            int a = (int)swDocTemplateTypes_e.swDocTemplateTypeNONE;
            int b = (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel;
            IModelDoc2 model = SolidWorks.OpenDoc6(path, a, b, string.Empty, ref qe, ref qw);
            Debug.WriteLine("model {0} loaded in mememory, full path {1} ... ok", model.GetTitle(), model.GetPathName());

            return model;
        }

        public IComponent2 AddModelToAssembly(string componentName, double x, double y, double z)
        {
            int a = (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig;
            IComponent2 component = Assembly.AddComponent5(componentName, a, string.Empty, false, string.Empty, x, y, z);
            Debug.WriteLine("model {0} added to assembly ... ok", componentName, string.Empty);

            return component;
        }
    }
}
