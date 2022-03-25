using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignObjectsToGroup: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignObjectsToGroup()
          : base("Assign > Objects to Groups", "AssgnObjctsGrp",
            "ETABSv1.SapModel...SetGroupAssign()\nAdds or removes objects from a specified group. ",
            "Tomaso", "06_Assign")
        {
        }
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.primary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_GroupName;
        int In_Remove;
        int In_ItemType;
        int In_ObjectType;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing object or group depending on the value of the ItemType item.", 
                GH_ParamAccess.list);
            In_GroupName = pManager.AddTextParameter("GroupName", "G", "The name of an existing group to which the assignment is made.",
                GH_ParamAccess.item);
            In_Remove = pManager.AddBooleanParameter("Remove", "R", "",
                GH_ParamAccess.item, false);
            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0 \nGroup = 1 \nSelectedObjects = 2 \nIf this item is Objects, the restraint assignment is made to the point object specified by the Name item. \nIf this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item. \nIf this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.",
                GH_ParamAccess.item, 0);
            In_ObjectType = pManager.AddTextParameter("ObjectType", "O", "'Point', 'Frame', 'Link', 'Area', 'Tendon'",
                GH_ParamAccess.item);

            pManager[In_Remove].Optional = true;
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
            string GroupName = string.Empty;
            bool Remove = false;
            int ItemType = 0;
            string ObjectType = string.Empty;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetData(In_GroupName, ref GroupName)) { return; }
            DA.GetData(In_Remove, ref Remove);
            DA.GetData(In_ItemType, ref ItemType);
            DA.GetData(In_ObjectType, ref ObjectType);


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
                    if (ObjectType == "Point")
                    {
                        ret = mySapModel.PointObj.SetGroupAssign(Name[i], GroupName, Remove, itemType);
                    }
                    if (ObjectType == "Frame")
                    {
                        ret = mySapModel.FrameObj.SetGroupAssign(Name[i], GroupName, Remove, itemType);
                    }
                    if (ObjectType == "Link")
                    {
                        ret = mySapModel.LinkObj.SetGroupAssign(Name[i], GroupName, Remove, itemType);
                    }
                    if (ObjectType == "Area")
                    {
                        ret = mySapModel.AreaObj.SetGroupAssign(Name[i], GroupName, Remove, itemType);
                    }
                    if (ObjectType == "Tendon")
                    {
                        ret = mySapModel.TendonObj.SetGroupAssign(Name[i], GroupName, Remove, itemType);
                    }
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.DGroup; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("fd0751d7-6360-4fe6-b255-b2b014bd31e9");
    }
}