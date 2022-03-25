using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{

    public class RunAnalysis: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public RunAnalysis()
          : base("Run Analysis", "RnAnlyss",
            "ETABSv1.Analyze.RunAnalysis()\nRuns the analysis. \nThe analysis model is automatically created as part of this function. IMPORTANT NOTE: Your model must have a file path defined before running the analysis.If the model is opened from an existing file, a file path is defined.If the model is created from scratch, the File.Save function must be called with a file name before running the analysis.Saving the file creates the file path.",
            "Tomaso", "07_Analyze")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary; }
        }

        //Initialize In and Out Parameters
        int In_Run;

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
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool Run = false;
            string iFilePath = string.Empty;

            if (!DA.GetData(In_Run, ref Run)) { return; }

            if (Run)
            { 
                ETABSv1.cOAPI myETABSObject = null;
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
                //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero)
                int ret = -1;

                //create an instance of the ETABS object from the latest installed ETABS
                try
                {
                    //create ETABS object
                    myETABSObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");
                }
                catch (Exception)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Cannot start a new instance of the program.");
                    return;
                }

                //start ETABS application
                ret = myETABSObject.ApplicationStart();

                //Get a reference to cSapModel to access all API classes and functions
                ETABSv1.cSapModel mySapModel = default(ETABSv1.cSapModel);
                mySapModel = myETABSObject.SapModel;

                //////////////////////////////////////////////////

                ret = mySapModel.Analyze.RunAnalysis();

                //////////////////////////////////////////////////
                
                //Check ret value
                if (ret != 0)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "API script FAILED to complete.");
                }
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
        public override Guid ComponentGuid => new Guid("e6c63c29-d05f-4777-9b58-3c5881ac404d");
    }
}