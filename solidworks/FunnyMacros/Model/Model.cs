using FunnyMacros.Util;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FunnyMacros.Model
{
    class Model
    {
        private static Converter<object, IFace2> CONVERTER = new Converter<object, IFace2>((o) => { return o as IFace2; });

        public Model(IMathUtility mathUtility)
        {
            MathUtility = mathUtility;
            Helper = Helper.Instance;//fucking hard code!
            //TODO: replace on more better solution
        }

        public Helper Helper { get; set; }
        public IMathUtility MathUtility { set; get; }
        public IComponent2 Component { get; set; }

        /// <summary>
        /// Box in local coordinates system
        /// </summary>
        public Box LBox
        {
            get { return new Box(Component.GetBox(false, false)); }
        }

        /// <summary>
        /// Box in global coordinates system
        /// </summary>
        public Box GBox
        {
            get { return new Box(Helper.ApplyTransform(Transform, LBox.CoordinatesCorners)); }
        }

        public Vector CenterLBox
        {
            get { return Helper.CenterBox(LBox); }
        }

        public Vector CenterGBox
        {
            get { return Helper.CenterBox(GBox); }
        }
        
        public MathTransform Transform
        {
            get { return Component.Transform2; }
            set { Component.Transform2 = value; }
        }

        public IBody2 Body
        {
            get { return Component.GetBody(); }
        }

        
        public IFace2[] Faces
        {
            get {  return Array.ConvertAll(Body.GetFaces(), CONVERTER); }
        }

        public void Translate(Vector translate)
        {
            IMathTransform transform = MathUtility.CreateTransform(null);
            double[] matrix = transform.ArrayData as double[];
            matrix[9] = translate.X;
            matrix[10] = translate.Y;
            matrix[11] = translate.Z;
            Transform = MathUtility.CreateTransform(matrix);
        }
    }
}
