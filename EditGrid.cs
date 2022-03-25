using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class EditGrid: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public EditGrid()
          : base("Edit > Grid", "EdtGrd",
            "ETABSv1.SapModel.DatabaseTables.SetTableForEditingArray()\nEdit Grid",
            "Tomaso", "02_Edit")
        {
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_XValue;
        int In_XLabel;
        int In_XBubbleLoc;
        int In_XBubbleVis;
        int In_YValue;
        int In_YLabel;
        int In_YBubbleLoc;
        int In_YBubbleVis;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "", 
                GH_ParamAccess.item, "G1");
            In_XValue = pManager.AddTextParameter("XValue", "XV", "", 
                GH_ParamAccess.list);
            In_XLabel = pManager.AddTextParameter("XLabel", "XL", "", 
                GH_ParamAccess.list);
            In_XBubbleLoc = pManager.AddTextParameter("XBubbleLocation", "XBL", "Start or End",
                GH_ParamAccess.item, "End");
            In_XBubbleVis = pManager.AddTextParameter("XBubbleVisible", "XBV", "XBubbleVisible",
                GH_ParamAccess.item, "Yes");
            In_YValue = pManager.AddTextParameter("YValue", "YV", "YValue",
                GH_ParamAccess.list);
            In_YLabel = pManager.AddTextParameter("YLabel", "YL", "YLabel",
                GH_ParamAccess.list);
            In_YBubbleLoc = pManager.AddTextParameter("YBubbleLocation", "YBL", "YBubbleLocation",
                GH_ParamAccess.item, "Start");
            In_YBubbleVis = pManager.AddTextParameter("YBubbleVisible", "YBV", "YBubbleVisible",
                GH_ParamAccess.item, "Yes");

            pManager[In_Name].Optional = true;
            pManager[In_XBubbleLoc].Optional = true;
            pManager[In_XBubbleVis].Optional = true;
            pManager[In_YBubbleLoc].Optional = true;
            pManager[In_YBubbleVis].Optional = true;
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
            bool iRun = false;
            string iName = "G1";
            List<string> iXValue = new List<String>();
            List<string> iXLabel = new List<String>();
            string iXBubbleLocation = "End";
            string iXBubbleVisible = "Yes";
            List<string> iYValue = new List<String>();
            List<string> iYLabel = new List<String>();
            string iYBubbleLocation = "Start";
            string iYBubbleVisible = "Yes";

            if (!DA.GetData(In_Run, ref iRun)) { return; }
            DA.GetData(In_Name, ref iName);
            if (!DA.GetDataList(In_XValue, iXValue)) { return; }
            if (!DA.GetDataList(In_XLabel, iXLabel)) { return; }
            DA.GetData(In_YBubbleLoc, ref iXBubbleLocation);
            DA.GetData(In_XBubbleVis, ref iXBubbleVisible);
            if (!DA.GetDataList(In_YValue, iYValue)) { return; }
            if (!DA.GetDataList(In_YLabel, iYLabel)) { return; }
            DA.GetData(In_YBubbleLoc, ref iYBubbleLocation);
            DA.GetData(In_YBubbleVis, ref iYBubbleVisible);

            //Test matching number of inputs
            if (iXValue.Count != iXLabel.Count || iYValue.Count != iYLabel.Count)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Value.Count != Label.Count");
                return;
            }

            string TableKey = "Grid Definitions - Grid Lines";

            if (iRun)
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
                string[] FieldsKeysIncluded = { "Name", "Grid Line Type", "ID", "Ordinate", "Angle", "X1", "Y1", "X2", "Y2", "Bubble Location", "Visible" };
                int NumberRecords = iXValue.Count + iYValue.Count;
                List<string> data = new List<string>();

                for (int i = 0; i < iXValue.Count; i++)
                {
                    string[] xdata = { iName, "X (Cartesian)", iXLabel[i], iXValue[i], "", "", "", "", "", iXBubbleLocation, iXBubbleVisible };
                    List<string> Xdata = new List<string>(xdata);
                    data.AddRange(Xdata);
                }
                for (int i = 0; i < iYValue.Count; i++)
                {
                    string[] ydata = { iName, "Y (Cartesian)", iYLabel[i], iYValue[i], "", "", "", "", "", iYBubbleLocation, iYBubbleVisible };
                    List<string> Ydata = new List<string>(ydata);
                    data.AddRange(Ydata);
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
        }


        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// You can add image files to your project resources and access them like this:
        /// return Resources.IconForThisComponent;
        /// </summary>

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.Edit; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("1c1e87a8-e9b7-4a36-9400-40bf16ed4de4");
    }
}