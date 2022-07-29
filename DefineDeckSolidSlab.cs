using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DefineDeckSolidSlab: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DefineDeckSolidSlab()
          : base("Define > Deck > Solid Slab", "DfnDckSldSlb",
            "ETABSv1.SapModel.PropArea.SetDeckSolidSLab()\nInitializes a solid-slab deck property",
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
        int In_SlabDepth;
        int In_ShearStudDia;
        int In_ShearStudHt;
        int In_ShearStudFu;

        int Out_Ret;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing deck property created with the Define Deck Section component.", 
                GH_ParamAccess.list);
            In_SlabDepth = pManager.AddNumberParameter("SlabDepth", "S", "Slab Depth, tc ",
                GH_ParamAccess.list);
            In_ShearStudDia = pManager.AddNumberParameter("ShearStudDia", "Sd", "Shear Stud Diameter ",
                GH_ParamAccess.list);
            In_ShearStudHt = pManager.AddNumberParameter("ShearStudHt", "Sd", "Shear Stud Height, hs ",
                GH_ParamAccess.list);
            In_ShearStudFu = pManager.AddNumberParameter("ShearStudFu", "Sf", "Shear Stud Tensile Strength, Fu ",
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
            List<double> SlabDepth = new List<double>();
            List<double> ShearStudDia = new List<double>();
            List<double> ShearStudHt = new List<double>();
            List<double> ShearStudFu = new List<double>();

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetDataList(In_SlabDepth, SlabDepth)) { return; }
            if (!DA.GetDataList(In_ShearStudDia, ShearStudDia)) { return; }
            if (!DA.GetDataList(In_ShearStudHt, ShearStudHt)) { return; }
            if (!DA.GetDataList(In_ShearStudFu, ShearStudFu)) { return; }

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

                if (SlabDepth.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        SlabDepth.Add(SlabDepth[0]);
                    }
                }

                if (ShearStudDia.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        ShearStudDia.Add(ShearStudDia[0]);
                    }
                }

                if (ShearStudHt.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        ShearStudHt.Add(ShearStudHt[0]);
                    }
                }

                if (ShearStudFu.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        ShearStudFu.Add(ShearStudFu[0]);
                    }
                }
                for (int i = 0; i < Name.Count; i++)
                {
                    ret = mySapModel.PropArea.SetDeckSolidSlab(Name[i], SlabDepth[i], ShearStudDia[i], ShearStudHt[i], ShearStudFu[i]);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.DDeck; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("abf0a70b-81b8-4d9b-a151-4ffd239674a5");
    }
}