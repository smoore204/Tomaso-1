using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignJointPanelZone: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignJointPanelZone()
          : base("Assign > Joint > Panel Zone", "AssgnJntPnlZn",
            "ETABSv1.SapModel.PointObj.SetPanelZone()",
            "Tomaso", "06_Assign")
        {
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_PropType;
        int In_Thickness;
        int In_K1;
        int In_K2;
        int In_LinkProp;
        int In_Connectivity;
        int In_LocalAxisFrom;
        int In_LocalAxisAngle;
        int In_ItemType;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "",
                GH_ParamAccess.list);
            In_PropType = pManager.AddIntegerParameter("PropType", "P", "",
                GH_ParamAccess.item);
            In_Thickness = pManager.AddNumberParameter("Thickness", "T", "",
                GH_ParamAccess.item);
            In_K1 = pManager.AddNumberParameter("K1", "K1", "",
                GH_ParamAccess.item);
            In_K2 = pManager.AddNumberParameter("K2", "K2", "",
                GH_ParamAccess.item);
            In_LinkProp = pManager.AddTextParameter("LinkProp", "L", "",
                GH_ParamAccess.item);
            In_Connectivity = pManager.AddIntegerParameter("Connectivity", "C", "",
                GH_ParamAccess.item);
            In_LocalAxisFrom = pManager.AddIntegerParameter("LocalAxisFrom", "F", "",
                GH_ParamAccess.item);
            In_LocalAxisAngle = pManager.AddNumberParameter("LocalAxisAngle", "A", "",
                GH_ParamAccess.item);
            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0 \nGroup = 1 \nSelectedObjects = 2 \nIf this item is Objects, the restraint assignment is made to the point object specified by the Name item. \nIf this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item. \nIf this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.",
                GH_ParamAccess.item);

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
            int PropType = -1;
            double Thickness = -1;
            double K1 = -1;
            double K2 = -1;
            string LinkProp = string.Empty;
            int Connectivity = -1;
            int LocalAxisFrom = -1;
            double LocalAxisAngle = -1;
            int ItemType = 0;

            if (!DA.GetData(In_Run, ref Run)) { return; }
            if (!DA.GetDataList(In_Name, Name)) { return; }
            if (!DA.GetData(In_PropType, ref PropType)) { return; }
            if (!DA.GetData(In_Thickness, ref Thickness)) { return; }
            if (!DA.GetData(In_K1, ref K1)) { return; }
            if (!DA.GetData(In_K2, ref K2)) { return; }
            if (!DA.GetData(In_LinkProp, ref LinkProp)) { return; }
            if (!DA.GetData(In_Connectivity, ref Connectivity)) { return; }
            if (!DA.GetData(In_LocalAxisFrom, ref LocalAxisFrom)) { return; }
            if (!DA.GetData(In_LocalAxisAngle, ref LocalAxisAngle)) { return; }
            if (!DA.GetData(In_ItemType, ref ItemType)) { return; }

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
                
                eItemType Type = eItemType.Objects;

                switch (ItemType)
                {
                    case 0:
                        Type = eItemType.Objects;
                        break;
                    case 1:
                        Type = eItemType.SelectedObjects;
                        break;
                    case 2:
                        Type = eItemType.Group;
                        break;
                }

                for (int i = 0; i < Name.Count; i++)
                {
                    ret = mySapModel.PointObj.SetPanelZone(Name[i], PropType, Thickness, K1, K2, LinkProp, Connectivity, LocalAxisFrom, LocalAxisAngle, Type);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AJointPanelZone; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("565da87a-adb2-47b8-a74d-d369a008295b");
    }
}