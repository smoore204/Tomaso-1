using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{

    public class SaveModel: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public SaveModel()
          : base("Save", "Save",
            "ETABSv1.SapModel.File.Save()Saves the present model.",
            "Tomaso", "01_File")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.quarternary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_FileName;
        int In_FilePath;
        int In_Close;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_FileName = pManager.AddTextParameter("FileName", "N", "",
                GH_ParamAccess.item);
            In_FilePath = pManager.AddTextParameter("FilePath", "P", "",
                GH_ParamAccess.item);
            In_Close = pManager.AddBooleanParameter("Close", "C", "Close instance after saving",
                GH_ParamAccess.item, false);

            pManager[In_Close].Optional = true;
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
            string FileName = string.Empty;
            string FilePath = string.Empty;
            bool Close = false;

            if (!DA.GetData(In_Run, ref Run)) { return; }
            if (!DA.GetData(In_FileName, ref FileName)) { return; }
            if (!DA.GetData(In_FilePath, ref FilePath)) { return; }
            DA.GetData(In_Close, ref Close);

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

                //////////////////////////////////////////////////

                string FullPath = FilePath + @"\" + FileName;
                ret = mySapModel.File.Save(FullPath);

                if (Close)
                {
                    ret = myETABSObject.ApplicationExit(true);
                }

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.Save; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("df26c193-2dca-4941-a3ce-5291909d7900");
    }
}