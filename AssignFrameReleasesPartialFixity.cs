using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignFrameReleasesPartialFixity : GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignFrameReleasesPartialFixity()
          : base("Assign > Frame > Releases/Partial Fixity", "AssgnFrmRlssPrtlFxty",
            "ETABSv1.SapModel.FrameObj.SetReleases()\nMakes end release and partial fixity assignments to frame objects. \nPartial fixity assignments are made to degrees of freedom that have been released only. \nSome release assignments would cause instability in the model.An error is returned if this type of assignment is made.Unstable release assignments include the following:\nU1 released at both ends\nU2 released at both ends\nU3 released at both ends\nR1 released at both ends\nR2 released at both ends and U3 at either end\nR3 released at both ends and U2 at either end", "Tomaso",
            "06_Assign")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_II;
        int In_JJ;
        int In_StartValue;
        int In_EndValue;
        int In_ItemType;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);

            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing frame object or group, depending on the value of the ItemType item.", 
                GH_ParamAccess.list);

            In_II = pManager.AddBooleanParameter("II", "I", "This is a list of six booleans indicating the I-End releases for the frame object. \nii(0) = U1 release; \nii(1) = U2 release \nii(2) = U3 release; \nii(3) = R1 release; \nii(4) = R2 release; \nii(5) = R3 release.",
                GH_ParamAccess.list, new List<bool> { false, false, false, false, false, false });

            In_JJ = pManager.AddBooleanParameter("JJ", "J", "This is a list of six booleans indicating the J-End releases for the frame object. \njj(0) = U1 release; \njj(1) = U2 release \njj(2) = U3 release; \njj(3) = R1 release; \njj(4) = R2 release; \njj(5) = R3 release.",
                GH_ParamAccess.list, new List<bool> { false, false, false, false, false, false });

            In_StartValue = pManager.AddNumberParameter("StartValue", "S", "This is a list of six values indicating the I-End partial fixity springs for the frame object. \nStartValue(0) = U1 partial fixity [F/L]; \nStartValue(1) = U2 partial fixity [F/L]; \nStartValue(2) = U3 partial fixity [F/L]; \nStartValue(3) = R1 partial fixity [FL/rad]; \nStartValue(4) = R2 partial fixity [FL/rad]; \nStartValue(5) = R3 partial fixity [FL/rad].",
                GH_ParamAccess.list, new List<double> { 0, 0, 0, 0, 0, 0 });

            In_EndValue = pManager.AddNumberParameter("EndValue", "E", "This is a list of six values indicating the J-End partial fixity springs for the frame object. \nStartValue(0) = U1 partial fixity [F/L]; \nStartValue(1) = U2 partial fixity [F/L]; \nStartValue(2) = U3 partial fixity [F/L]; \nStartValue(3) = R1 partial fixity [FL/rad]; \nStartValue(4) = R2 partial fixity [FL/rad]; \nStartValue(5) = R3 partial fixity [FL/rad].",
                GH_ParamAccess.list, new List<double> { 0, 0, 0, 0, 0, 0 });

            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0; \nGroup = 1; \nSelectedObjects = 2. \nIf this item is Objects, the assignment is made to the frame object specified by the Name item. \nIf this item is Group, the assignment is made to all frame objects in the group specified by the Name item. \nIf this item is SelectedObjects, the assignment is made to all selected frame objects and the Name item is ignored.",
                GH_ParamAccess.item, 0) ;

            pManager[In_II].Optional = true;
            pManager[In_JJ].Optional = true;
            pManager[In_StartValue].Optional = true;
            pManager[In_EndValue].Optional = true;
            pManager[In_ItemType].Optional = true;
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
            List<bool> II = new List<bool>();
            List<bool> JJ = new List<bool>();
            List<double> StartValue = new List<double>();
            List<double> EndValue = new List<double>();
            int ItemType = -1;

            if (!DA.GetData(In_Run, ref Run)) { return; }
            if (!DA.GetDataList(In_Name, Name)) { return; }
            DA.GetDataList(In_II, II);
            DA.GetDataList(In_JJ, JJ);
            DA.GetDataList(In_StartValue, StartValue);
            DA.GetDataList(In_EndValue, EndValue);
            DA.GetData(In_ItemType, ref ItemType);

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

                bool[] ii = II.ToArray();
                bool[] jj = JJ.ToArray();
                double[] startValue = StartValue.ToArray();
                double[] endValue = EndValue.ToArray();

                eItemType itemType = eItemType.Objects;

                switch (ItemType)
                {
                    case 0:
                        itemType = eItemType.Objects;
                        break;
                    case 1:
                        itemType = eItemType.Group;
                        break;
                    case 2:
                        itemType = eItemType.SelectedObjects;
                        break;
                }

                for (int i = 0; i < Name.Count; i++)
                {
                    ret = mySapModel.FrameObj.SetReleases(Name[i], ref ii, ref jj, ref startValue, ref endValue, itemType);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AFrameRelease; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("a2643218-8260-4be6-98c8-0964458b59e2");
    }
}