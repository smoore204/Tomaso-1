using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DesignSteelSetGroup: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DesignSteelSetGroup()
          : base("Design > Steel Frame Design > Select Design Group", "DsgnStlFrmDsgnGrp",
            "ETABSv1.SapModel.DesignSteel.SetGroup()\n",
            "Tomaso", "08_Design")
        {
        }
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.primary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_Selected;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            
            In_Name = pManager.AddTextParameter("Name", "N", "Retrieves the design section for a specified steel frame object ", GH_ParamAccess.list);
            In_Selected = pManager.AddBooleanParameter("Selected", "S", "", GH_ParamAccess.item, true);

            pManager[In_Selected].Optional = true;
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
            bool Selected = true;

            if (!DA.GetData(In_Run, ref Run)) { return; }
            if (!DA.GetDataList(In_Name, Name)) { return; }
            DA.GetData(In_Selected, ref Selected);

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

                /////////////////////////////////////////////////

                for (int i = 0; i < Name.Count; i++)
                {
                    ret = mySapModel.DesignSteel.SetGroup(Name[i], Selected);

                    //////////////////////////////////////////////////

                    //Check ret value
                    if (ret != 0)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "API script FAILED to complete.");
                    }

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.Run; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("9710e5fd-d8b7-4fc0-b954-609f0797cef0");
    }
}