using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DefineFrameSectionsAutoSelectConcrete: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DefineFrameSectionsAutoSelectConcrete()
          : base("Define > Frame Sections > Auto Select Concrete", "DfnFrmSctnsAtSlctCncrt",
            "ETABSv1.SapModel.DatabaseTables.SetTableForEditingArray()\n",
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
        int In_DesignType;
        int In_FromFile;
        int In_FileName;
        int In_StartingSection;
        int In_IncludedSection;
        int In_GUID;
        int In_Notes;
        int Out_Ret;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "", 
                GH_ParamAccess.item);
            In_DesignType = pManager.AddTextParameter("DesignType", "D", "", 
                GH_ParamAccess.item, "Concrete");
            In_FromFile = pManager.AddTextParameter("FromFile", "F", "", 
                GH_ParamAccess.item, "No");
            In_FileName = pManager.AddTextParameter("FileName", "F", "",
                GH_ParamAccess.item);
            In_StartingSection = pManager.AddTextParameter("StartingSection", "S", "",
                GH_ParamAccess.item, "Median");
            In_IncludedSection = pManager.AddTextParameter("IncludedSection", "I", "",
                GH_ParamAccess.list);
            In_GUID = pManager.AddTextParameter("GUID", "G", "",
                GH_ParamAccess.item);
            In_Notes = pManager.AddTextParameter("Notes", "N", "",
                GH_ParamAccess.item);

            pManager[In_DesignType].Optional = true;
            pManager[In_FromFile].Optional = true;
            pManager[In_FileName].Optional = true;
            pManager[In_StartingSection].Optional = true;
            pManager[In_GUID].Optional = true;
            pManager[In_Notes].Optional = true;
        }


        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            Out_Ret = pManager.AddBooleanParameter("Ret", "R", "",
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
            string Name = "";
            string DesignType = "Concrete";
            string FromFile = "No";
            string FileName = "";
            string StartingSection = "Median";
            List<string> IncludedSection = new List<String>();
            string GUID = "";
            string Notes = "";

            if (!DA.GetData(In_Run, ref Run)) { return; }
            if (!DA.GetData(In_Name, ref Name)) { return; }
            DA.GetData(In_DesignType, ref DesignType);
            DA.GetData(In_FromFile, ref FromFile);
            DA.GetData(In_FileName, ref FileName);
            DA.GetData(In_StartingSection, ref StartingSection);
            if (!DA.GetDataList(In_IncludedSection, IncludedSection)) { return; }
            DA.GetData(In_GUID, ref GUID);
            DA.GetData(In_Notes, ref Notes);

            //Test matching number of inputs

            string TableKey = "Frame Section Property Definitions - Auto Select List";

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
                int TableVersion = 0;
                string[] FieldsKeysIncluded = { "Name", "DesignType", "From File?", "File Name", "Starting Section", "Included Section", "GUID", "Notes"};
                int NumberRecords = IncludedSection.Count;
                List<string> data = new List<string>();

                for (int i = 0; i < IncludedSection.Count; i++)
                {
                    string[] xdata = { Name, DesignType, FromFile, FileName, StartingSection, IncludedSection[i], GUID, Notes};
                    List<string> Xdata = new List<string>(xdata);
                    data.AddRange(Xdata);
                }

                string[] Data = data.ToArray();
                ret = mySapModel.DatabaseTables.SetTableForEditingArray(TableKey, ref TableVersion, ref FieldsKeysIncluded, NumberRecords, ref Data);

                bool FillImport = true;
                int NumFatalErrors = 0;
                int NumErrorMsgs = 0;
                int NumWarnMsgs = 0;
                int NumInfoMsgs = 0;
                string ImportLog = "";

                ret = mySapModel.DatabaseTables.ApplyEditedTables(FillImport, ref NumFatalErrors, ref NumErrorMsgs, ref NumWarnMsgs, ref NumInfoMsgs, ref ImportLog);

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
        public override Guid ComponentGuid => new Guid("7d7cfe9854bf486cbe0364514497d083");
    }
}