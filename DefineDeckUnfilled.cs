using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DefineDeckUnfilled: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DefineDeckUnfilled()
          : base("Define > Deck > Unfilled", "DfnDckUnflld",
            "ETABSv1.SapModel.PropArea.SetDeckUnfilled()\nInitializes an unfilled deck property",
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
        int In_RibDepth;
        int In_RibWidthTop;
        int In_RibWidthBot;
        int In_RibSpacing;
        int In_ShearThickness;
        int In_UnitWeight;


        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing deck property created with the Define Deck Section component.", 
                GH_ParamAccess.list);
            In_RibDepth = pManager.AddNumberParameter("RibDepth", "Rd", "Rib Depth, hr ",
                GH_ParamAccess.list);
            In_RibWidthTop = pManager.AddNumberParameter("RibidthTop", "Rt", "Rib Width Top, wrt",
                GH_ParamAccess.list);
            In_RibWidthBot = pManager.AddNumberParameter("RibWidthBot", "Rb", "Rib Width Bottom, wrb",
                GH_ParamAccess.list);
            In_RibSpacing = pManager.AddNumberParameter("RibSpacing", "Rs", "Rib Spacing, sr",
                GH_ParamAccess.list);
            In_ShearThickness = pManager.AddNumberParameter("ShearThickness", "S", "Deck Shear Thickness ",
                GH_ParamAccess.list);
            In_UnitWeight = pManager.AddNumberParameter("UnitWeight", "U", "Deck Unit Weight ",
                GH_ParamAccess.list);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
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
            List<double> SlabDepth = new List<double>();
            List<double> RibDepth = new List<double>();
            List<double> RibWidthTop = new List<double>();
            List<double> RibWidthBot = new List<double>();
            List<double> RibSpacing = new List<double>();
            List <double> ShearThickness = new List<double>();
            List<double> UnitWeight = new List<double>();
            List<double> ShearStudDia = new List<double>();
            List<double> ShearStudHt = new List<double>();
            List<double> ShearStudFu = new List<double>();

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetDataList(In_RibDepth, RibDepth)) { return; }
            if (!DA.GetDataList(In_RibWidthTop, RibWidthTop)) { return; }
            if (!DA.GetDataList(In_RibWidthBot, RibWidthBot)) { return; }
            if (!DA.GetDataList(In_RibSpacing, RibSpacing)) { return; }
            if (!DA.GetDataList(In_ShearThickness, ShearThickness)) { return; }
            if (!DA.GetDataList(In_UnitWeight, UnitWeight)) { return; }

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

                if (RibDepth.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        RibDepth.Add(RibDepth[0]);
                    }
                }
                if (RibWidthTop.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        RibWidthTop.Add(RibWidthTop[0]);
                    }
                }

                if (RibWidthBot.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        RibWidthBot.Add(RibWidthBot[0]);
                    }
                }
       
                if (RibSpacing.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        RibSpacing.Add(RibSpacing[0]);
                    }
                }

                if (ShearThickness.Count == 1)
                {
                    for (int i = 0; i < Name.Count-1; i++)
                    {
                        ShearThickness.Add(ShearThickness[0]);
                    }
                }

                if (UnitWeight.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        UnitWeight.Add(UnitWeight[0]);
                    }
                }

                for (int i = 0; i < Name.Count; i++)
                {
                    ret = mySapModel.PropArea.SetDeckUnfilled(Name[i], RibDepth[i], RibWidthTop[i], RibWidthBot[i], RibSpacing[i], ShearThickness[i], UnitWeight[i]);
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
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// You can add image files to your project resources and access them like this:
        /// return Resources.IconForThisComponent;
        /// </summary>

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.DDeck; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("1ace0e89-9be5-4776-9531-3711a8a6b82c");
    }
}