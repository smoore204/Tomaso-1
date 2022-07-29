using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DefineSlabRibbed: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DefineSlabRibbed()
          : base("Define > Slab > Ribbed", "DfnSlbRbbd",
            "ETABSv1.SapModel.PropArea.SetSlabRibbed()\nInitializes a ribbed slab property",
            "Tomaso", "04_Define")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_OverallDepth;
        int In_SlabThickness;
        int In_StemWidthTop;
        int In_StemWidthBot;
        int In_RibSpacing;
        int In_RibsParallelTo;

        int Out_Ret;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing ribbed slab property. ", 
                GH_ParamAccess.list);
            In_OverallDepth = pManager.AddNumberParameter("OverallDepth", "O", "Overall Depth",
                GH_ParamAccess.list);
            In_SlabThickness = pManager.AddNumberParameter("SlabThickness", "S", "Slab Thickness ",
                GH_ParamAccess.list);
            In_StemWidthTop = pManager.AddNumberParameter("StemWidthTop", "St", "Stem Width at Top",
                GH_ParamAccess.list);
            In_StemWidthBot = pManager.AddNumberParameter("StemWidthBot", "Sb", "Stem Width at Bottom ",
                GH_ParamAccess.list);
            In_RibSpacing = pManager.AddNumberParameter("RibSpacing", "Rs", "Rib Spacing (Perpendicular to Rib Direction) ",
                GH_ParamAccess.list);
            In_RibsParallelTo = pManager.AddIntegerParameter("RibsParallelTo", "Rp", "This is the Local Axis that the Rib Direction is Parallel to \nThis value is 1 for the Local 1 Axis, and 2 for the Local 2 Axis.",
                GH_ParamAccess.list);

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            Out_Ret = pManager.AddIntegerParameter("Ret", "R", "",
                GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool Run = false;
            List<string> Name = new List<string>();
            List<double> OverallDepth = new List<double>();
            List<double> SlabThickness = new List<double>();
            List<double> StemWidthTop = new List<double>();
            List<double> StemWidthBot = new List<double>();
            List<double> RibSpacing  = new List<double>();
            List<int> RibsParallelTo  = new List<int>();

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetDataList(In_OverallDepth, OverallDepth)) { return; }
            if (!DA.GetDataList(In_SlabThickness, SlabThickness)) { return; }
            if (!DA.GetDataList(In_StemWidthTop, StemWidthTop)) { return; }
            if (!DA.GetDataList(In_StemWidthBot, StemWidthBot)) { return; }
            if (!DA.GetDataList(In_RibSpacing, RibSpacing )) { return; }
            if (!DA.GetDataList(In_RibsParallelTo, RibsParallelTo )) { return; }

            if (Run)
            {
                //Dimension the ETABS Object as cOAPI type
                ETABSv1.cOAPI myETABSObject = null;

                //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero)
                int ret = -1;

                //Create API helper object
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

                //Get the active ETABS object
                myETABSObject = myHelper.GetObject("CSI.ETABS.API.ETABSObject");
                ETABSv1.cSapModel mySapModel = default(ETABSv1.cSapModel);

                try
                {
                    mySapModel = myETABSObject.SapModel;
                }
                catch (Exception)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "No running instance of the program found or failed to attach.");
                    return;
                }
                //////////////////////////////////////////////////

                if (OverallDepth.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        OverallDepth.Add(OverallDepth[0]);
                    }
                }
                if (SlabThickness.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        SlabThickness.Add(SlabThickness[0]);
                    }
                }
                if (StemWidthTop.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        StemWidthTop.Add(StemWidthTop[0]);
                    }
                }

                if (StemWidthBot.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        StemWidthBot.Add(StemWidthBot[0]);
                    }
                }
       

                if (RibSpacing.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        RibSpacing.Add(RibSpacing [0]);
                    }
                }

                if (RibsParallelTo.Count == 1)
                {
                    for (int i = 0; i < Name.Count-1; i++)
                    {
                        RibsParallelTo.Add(RibsParallelTo[0]);
                    }
                }

                for (int i = 0; i < Name.Count; i++)
                {
                    ret = mySapModel.PropArea.SetSlabRibbed(Name[i], OverallDepth[i], SlabThickness[i], StemWidthTop[i], StemWidthBot[i], RibSpacing[i], RibsParallelTo[i]);
                }
                
                
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
            DA.SetData(Out_Ret, Run);
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// You can add image files to your project resources and access them like this:
        /// return Resources.IconForThisComponent;
        /// </summary>

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.DSlab; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("5b8ecc15-0d1f-4737-b012-704a9b6a1efe");
    }
}