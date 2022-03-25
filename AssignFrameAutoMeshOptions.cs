using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignFrameAutoMeshOptions: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignFrameAutoMeshOptions()
          : base("Assign > Frame > Frame Auto Mesh Options", "AssnFrameAtMshOptns",
            "ETABSv1.SapModel.DatabaseTables.SetTableForEditingArray()\nAssign Frame Auto Mesh Options",
            "Tomaso", "06_Assign")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary; }
        }

        //Initialize In and Out Parameters

        int In_Run;
        int In_AutoMesh;
        int In_AtIntermediateJoints;
        int In_AtIntersections;
        int In_MinNumber;
        int In_NumberSegments;
        int In_MaxLength;
        int In_SegmentLength;
        int In_Name;


        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "",
                GH_ParamAccess.list);
            In_AutoMesh = pManager.AddTextParameter("AutoMesh", "A", "'Yes' or 'No'", 
                GH_ParamAccess.item);
            In_AtIntermediateJoints = pManager.AddTextParameter("AtIntermediateJoints", "I", "'Yes' or 'No', not required if AutoMesh = No",
                GH_ParamAccess.item);
            In_AtIntersections = pManager.AddTextParameter("AtInteresections", "I", "'Yes' or 'No', not required if AutoMesh = No",
                GH_ParamAccess.item);
            In_MinNumber = pManager.AddTextParameter("MinNumber?", "M", "'Yes' or 'No', not required if AutoMesh = No",
                GH_ParamAccess.item);
            In_NumberSegments = pManager.AddTextParameter("NumberSegments", "N", "not required if AutoMesh = No",
                GH_ParamAccess.item);
            In_MaxLength = pManager.AddTextParameter("MaxLength?", "M", "'Yes' or 'No', not required if AutoMesh = No",
                GH_ParamAccess.item);
            In_SegmentLength = pManager.AddTextParameter("SegmentLengthft", "S", "Not required if AutoMesh = No",
                GH_ParamAccess.item);

            pManager[In_AtIntermediateJoints].Optional = true;
            pManager[In_AtIntersections].Optional = true;
            pManager[In_MinNumber].Optional = true;
            pManager[In_NumberSegments].Optional = true;
            pManager[In_MaxLength].Optional = true;
            pManager[In_SegmentLength].Optional = true;


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
            string AutoMesh = string.Empty;
            string AtIntermediateJoints = string.Empty;
            string AtIntersections = string.Empty;
            string MinNumber = string.Empty;
            string NumberSegments = string.Empty;
            string MaxLength = string.Empty;
            string SegmentLength = string.Empty;

            if (!DA.GetData(In_Run, ref Run)) { return; }
            if (!DA.GetDataList(In_Name, Name)) { return; }
            if (!DA.GetData(In_AutoMesh, ref AutoMesh)) { return; }
            DA.GetData(In_AtIntermediateJoints, ref AtIntermediateJoints);
            DA.GetData(In_AtIntersections, ref AtIntersections);
            DA.GetData(In_MinNumber, ref MinNumber);
            DA.GetData(In_NumberSegments, ref NumberSegments);
            DA.GetData(In_MaxLength, ref MaxLength);
            DA.GetData(In_SegmentLength, ref SegmentLength);

            //Test matching number of inputs

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
                ///
                int TableVersion = 0;
                string TableKey = "Frame Assignments - Frame Auto Mesh Options";
                string[] FieldsKeysIncluded = { "UniqueName", "Auto Mesh", "At Intermediate Joints", "At Intersections", "Min Number?", "Number Segments", "MaxLength?", "Segment Length ft" };
                int NumberRecords = Name.Count*8;
                List<string> Data = new List<string>();

                for (int i = 0; i < Name.Count; i++)
                {
                    if (AutoMesh == "No")
                    {
                        List<string> data = new List<string> { Name[i], AutoMesh, "", "", "", "", "", "" };
                        Data.AddRange(data);
                    }
                    else
                    {
                        List<string> data = new List<string> { Name[i], AutoMesh, AtIntermediateJoints, AtIntersections, MinNumber, NumberSegments, SegmentLength };
                        Data.AddRange(data);
                    }
                }

                string[] DataArr = Data.ToArray();
                ret = mySapModel.DatabaseTables.SetTableForEditingArray(TableKey, ref TableVersion, ref FieldsKeysIncluded, NumberRecords, ref DataArr);

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AssgnFrmAtMshOptns; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("e6ca6619-eceb-4daa-9c16-a92c7651c975");
    }
}