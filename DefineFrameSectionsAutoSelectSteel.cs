using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DefineFrameSectionsAutoSelectSteel: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DefineFrameSectionsAutoSelectSteel()
          : base("Define > Frame Sections > Auto Select Steel", "DfnFrmSctnsAtSlctSteel",
            "ETABSv1.SapModel.Propframe.SetAutoSelectSteel()\nRetrieves frame section property data for a steel auto select lists.",
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
        int In_NumberItems;
        int In_SectName;
        int In_AutoStartSection;

        int Out_Ret;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.", 
                GH_ParamAccess.list);
            In_NumberItems = pManager.AddIntegerParameter("NumberItems", "NI", "The number of frame section properties included in the auto select list.",
                GH_ParamAccess.item);
            In_SectName = pManager.AddTextParameter("SectName", "S", "This is an array of the names of the frame section properties included in the auto select list. Auto select lists and nonprismatic(variable) sections are not allowed in this array.",
                GH_ParamAccess.list);
            In_AutoStartSection = pManager.AddTextParameter("AutoStartSelection", "A", "This is either Median or the name of a frame section property in the SectName array. It is the starting section for the auto select list.",
                GH_ParamAccess.item, "Median");

            pManager[In_AutoStartSection].Optional = true;
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
            int NumberItems = -1;
            List<string> SectName = new List<string>();
            string AutoStartSection = "Median";
            string GUID = "";
            string Notes = "";

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetData(In_NumberItems, ref NumberItems)) { return; }
            if (!DA.GetDataList(In_SectName, SectName)) { return; }
            DA.GetData(In_AutoStartSection, ref AutoStartSection);


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
                string[] sectName = SectName.ToArray();

                for (int i = 0; i < Name.Count; i++)
                {
                    ret = mySapModel.PropFrame.SetAutoSelectSteel(Name[i], NumberItems, ref sectName, AutoStartSection, Notes, GUID);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AFrameSection; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("2bc61678-d8c5-4a1c-b564-862cadc5f158");
    }
}