using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace FunnyMacros.Util
{
    class LoadUtil
    {
        //quantity of errors
        private static int qe;

        //quantity of warnings
        private static int qw;

        public static IModelDoc2 LoadAssembly(ISldWorks application, string path)
        {
            int a = (int)swDocTemplateTypes_e.swDocTemplateTypePART;
            IModelDoc2 model = application.OpenDoc6(path, a, 0, string.Empty, ref qe, ref qw);
            Debug.WriteLine("assemly loaded. errors: {2}, warnings: {3}, pathName: {0}, title: {1} ... ok", model.GetPathName(), model.GetTitle(), qe, qw);

            application.ActivateDoc2(model.GetTitle(), false, ref qe);
            Debug.WriteLine("activate doc. errors: {0} ... ok", qe, string.Empty);

            model = application.ActiveDoc as IModelDoc2;
            Debug.WriteLine("set assembly to active document ... ok");

            return model;
        }

        public static IModelDoc2 LoadModel(ISldWorks solidWorks, string path)
        {
            int a = (int)swDocTemplateTypes_e.swDocTemplateTypeNONE;
            int b = (int)swOpenDocOptions_e.swOpenDocOptions_LoadModel;
            IModelDoc2 model = solidWorks.OpenDoc6(path, a, b, string.Empty, ref qe, ref qw);
            Debug.WriteLine("model {0} loaded in mememory, full path {1} ... ok", model.GetTitle(), model.GetPathName());

            return model;
        }

        public static IComponent2 AddModelToAssembly(IAssemblyDoc assembly, string componentName, double x, double y, double z)
        {
            int a = (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig;
            IComponent2 component = assembly.AddComponent5(componentName, a, string.Empty, false, string.Empty, x, y, z);
            Debug.WriteLine("model {0} added to assembly ... ok", componentName, string.Empty);

            return component;
        }
    }
}
