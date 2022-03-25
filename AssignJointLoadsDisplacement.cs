using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignJointLoadsDisplacement: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignJointLoadsDisplacement()
          : base("Assign > Joint Loads > Ground Displacement", "AssgnJntLdsGrndDsplcmnt",
            "ETABSv1.SapModel.PointObj.SetLoadDispl()\nMakes ground displacement load assignments to point objects.",
            "Tomaso", "06_Assign")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.quinary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_LoadPat;
        int In_Value;
        int In_Replace;
        int In_CSys;
        int In_ItemType;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing point object or group depending on the value of the ItemType item.", 
                GH_ParamAccess.list);
            In_LoadPat = pManager.AddTextParameter("LoadPat", "L", "The name of the load pattern for the point force load.",
                GH_ParamAccess.item);
            In_Value = pManager.AddNumberParameter("Value", "V", "This is a list of six point load values. \nValue(0) = U1[L] \nValue(1) = U2[L] \nValue(2) = U3[L] \nValue(3) = R1[rad] \nValue(4) = R2[rad] \nValue(5) = R3[rad]",
                GH_ParamAccess.list);
            In_Replace = pManager.AddBooleanParameter("Replace", "R", "If this item is True, all previous point force loads, if any, assigned to the specified point object(s) in the specified load pattern are deleted before making the new assignment. ",
                GH_ParamAccess.item, false);
            In_CSys = pManager.AddTextParameter("CSys", "C", "The name of the coordinate system for the considered point force load. This is Local or the name of a defined coordinate system.",
                GH_ParamAccess.item, "Global");
            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0 \nGroup = 1 \nSelectedObjects = 2 \nIf this item is Objects, the restraint assignment is made to the point object specified by the Name item. \nIf this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item. \nIf this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.",
                GH_ParamAccess.item, 0);

            pManager[In_Replace].Optional = true;
            pManager[In_CSys].Optional = true;
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
            string LoadPat = string.Empty;
            List<double> Value = new List<double>();
            bool Replace = false;
            string CSys = "Global";
            int ItemType = 0;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetData(In_LoadPat, ref LoadPat)) { return; }
            if (!DA.GetDataList(In_Value, Value)) { return; }
            DA.GetData(In_Replace, ref Replace);
            DA.GetData(In_CSys, ref CSys);
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

                eItemType itemType = eItemType.Objects;
                double[] value = Value.ToArray();

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
                    ret = mySapModel.PointObj.SetLoadDispl(Name[i], LoadPat, ref value, Replace, CSys, itemType);

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AJointLoadDispl; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("f6b97cc5-333d-4754-a8d0-8a12b3e429e7");
    }
}