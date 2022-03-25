using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class GetFrameSections: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public GetFrameSections()
          : base("Get > Frame Sections", "GtFrmSctns",
            "ETABSv1.SapModel.PropFrame.GetAllFrameProperties_2()\nRetrieves select data for all frame properties in the model.",
            "Tomaso", "04_Define")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.tertiary; }
        }

        //Initialize In and Out Parameters
        int In_Run;

        int Out_NumberNames;
        int Out_MyName;
        int Out_PropType;
        int Out_t3;
        int Out_t2;
        int Out_tf;
        int Out_tw;
        int Out_t2b;
        int Out_tfb;
        int Out_Area;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            Out_NumberNames = pManager.AddIntegerParameter("NumberNames", "N", "", GH_ParamAccess.item);
            Out_MyName = pManager.AddTextParameter("MyName", "M", "", GH_ParamAccess.list);
            Out_PropType = pManager.AddGenericParameter("PropType", "P", "", GH_ParamAccess.list);
            Out_t3 = pManager.AddNumberParameter("t3", "t3", "", GH_ParamAccess.list);
            Out_t2 = pManager.AddNumberParameter("t2", "t2", "", GH_ParamAccess.list);
            Out_tf = pManager.AddNumberParameter("tf", "tf", "", GH_ParamAccess.list);
            Out_tw = pManager.AddNumberParameter("tw", "tw", "", GH_ParamAccess.list);
            Out_t2b = pManager.AddNumberParameter("t2b", "t2b", "", GH_ParamAccess.list);
            Out_tfb = pManager.AddNumberParameter("tfb", "tfb", "", GH_ParamAccess.list);
            Out_Area = pManager.AddNumberParameter("Area", "A", "", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool Run = false;

            DA.GetData(In_Run, ref Run);

            if (Run)
            {
                //dimension the ETABS Object as cOAPI type
                ETABSv1.cOAPI myETABSObject = null;

                //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero)
                int ret = -1;

                //create API helper object
                ETABSv1.cHelper myHelper;
                try
                {
                    myHelper = new ETABSv1.Helper();
                }
                catch (Exception)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Cannot create an instance of the Helper object");
                    return;
                }

                //get the active ETABS object
                myETABSObject = myHelper.GetObject("CSI.ETABS.API.ETABSObject");
                ETABSv1.cSapModel mySapModel = default(ETABSv1.cSapModel);

                try
                {
                    mySapModel = myETABSObject.SapModel;
                }
                catch (Exception)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Could not find instance of Etabs to attach to");
                    return;
                }

                int NumberNames = 0;
                string[] MyName = new string[] { };
                ETABSv1.eFramePropType[] PropType = new ETABSv1.eFramePropType[] { };
                double[] t3 = new double[] { };
                double[] t2 = new double[] { };
                double[] tf = new double[] { };
                double[] tw = new double[] { };
                double[] t2b = new double[] { };
                double[] tfb = new double[] { };
                double[] Area = new double[] { };

                /////////////////////////////////////////////////

                ret = mySapModel.PropFrame.GetAllFrameProperties_2(ref NumberNames, ref MyName, ref PropType, ref t3, ref t2, 
                    ref tf, ref tw, ref t2b, ref tfb, ref Area);
                
                DA.SetData(Out_NumberNames, NumberNames);
                DA.SetDataList(Out_MyName, MyName);
                DA.SetDataList(Out_PropType, PropType);
                DA.SetDataList(Out_t3, t3);
                DA.SetDataList(Out_t2, t2);
                DA.SetDataList(Out_tf, tf);
                DA.SetDataList(Out_tw, tw);
                DA.SetDataList(Out_t2b, t2b);
                DA.SetDataList(Out_tfb, tfb);
                DA.SetDataList(Out_Area, Area);

                //////////////////////////////////////////////////

                //Check ret value
                if (ret != 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "API script FAILED to complete.");
                }

                //Clean up variables
                mySapModel = null;
                myETABSObject = null;
                
            }
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// You can add image files to your project resources and access them like this:
        /// return Resources.IconForThisComponent;
        /// </summary>

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.SophieMemoji; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("321935ed-36c8-4abd-b501-0030a7bf7d82");
    }
}