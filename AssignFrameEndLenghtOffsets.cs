using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignFrameEndLenghtOffsets: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignFrameEndLenghtOffsets()
          : base("Assign > Frame > End Length Offsets", "AssgnFrmEndLngthOffsts",
            "ETABSv1.SapModel.FrameObj.SetInsertionPoint()\nAssigns frame object end offsets along the 1-axis of the object.",
            "Tomaso", "06_Assign")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_AutoOffset;
        int In_Length1;
        int In_Length2;
        int In_RZ;
        int In_ItemType;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing frame object or group depending on the value of the ItemType item.", 
                GH_ParamAccess.list);
            In_AutoOffset = pManager.AddBooleanParameter("AutoOffset", "A", "If this item is True, the end length offsets are automatically determined by the program from object connectivity, and the Length1, Length2 and RZ items are ignored. ",
                GH_ParamAccess.item, false);
            In_Length1 = pManager.AddNumberParameter("Length1", "L1", "The offset length along the 1-axis of the frame object at the I-End of the frame object. [L] ",
                GH_ParamAccess.list, 0);
            In_Length2 = pManager.AddNumberParameter("Length2", "L2", "The offset length along the 1-axis of the frame object at the J-End of the frame object. [L] ",
                GH_ParamAccess.list, 0);
            In_RZ = pManager.AddNumberParameter("Rz", "R", "The rigid zone factor. This is the fraction of the end offset length assumed to be rigid for bending and shear deformations. ",
                GH_ParamAccess.list, 0);
            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0 \nGroup = 1 \nSelectedObjects = 2 \nIf this item is Objects, the restraint assignment is made to the point object specified by the Name item. \nIf this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item. \nIf this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.",
                GH_ParamAccess.item, 0);

            pManager[In_Length1].Optional = true;
            pManager[In_Length2].Optional = true;
            pManager[In_RZ].Optional = true;
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
            bool AutoOffset = false;
            List<double> Length1 = new List<double>();
            List<double> Length2 = new List<double>();
            List<double> RZ = new List<double>();
            int ItemType = 0;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetData(In_AutoOffset, ref AutoOffset)) { return; }
            if (!DA.GetDataList(In_Length1, Length1)) { return; }
            if (!DA.GetDataList(In_Length2, Length2)) { return; }
            if (!DA.GetDataList(In_RZ, RZ)) { return; }
            DA.GetData(In_ItemType, ref ItemType);

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
                if (Length1.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Length1.Add(Length1[0]);
                    }
                }

                if (Length2.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Length2.Add(Length2[0]);
                    }
                }

                if (RZ.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        RZ.Add(RZ[0]);
                    }
                }


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
                    ret = mySapModel.FrameObj.SetEndLengthOffset(Name[i], AutoOffset, Length1[i], Length2[i], RZ[i], itemType);

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AFrameELO; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("abe6fcfa-b0b6-453d-ac14-abef811eccf9");
    }
}