using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignFrameLoadsTemperature: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignFrameLoadsTemperature()
          : base("Assign > Frame Loads > Temperature", "AssgnFrmLdsTmprtr",
            "ETABSv1.SapModel.FrameObj.SetLoadTemperature()\nAssigns temperature loads to frame objects. ",
            "Tomaso", "06_Assign")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.senary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_LoadPat;
        int In_MyType;
        int In_Value;
        int In_PatternName;
        int In_Replace;
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
            In_LoadPat = pManager.AddTextParameter("LoadPat", "L", "The name of the load pattern.",
                GH_ParamAccess.list);
            In_MyType = pManager.AddIntegerParameter("Type", "T", "",
                GH_ParamAccess.list);
            In_Value = pManager.AddNumberParameter("Value", "V", "",
                GH_ParamAccess.list);
            In_PatternName = pManager.AddTextParameter("PatternName", "P", "",
                GH_ParamAccess.list, "");
            In_Replace = pManager.AddBooleanParameter("Replace", "R", "If this item is True, all previous loads, if any, assigned to the specified point object(s) in the specified load pattern are deleted before making the new assignment. ",
                GH_ParamAccess.list, true);
            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0 \nGroup = 1 \nSelectedObjects = 2 \nIf this item is Objects, the restraint assignment is made to the point object specified by the Name item. \nIf this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item. \nIf this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.",
                GH_ParamAccess.item, 0);

            pManager[In_PatternName].Optional = true;
            pManager[In_Replace].Optional = true;
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
            List<string> LoadPat = new List<string>();
            List<int> MyType = new List<int>();
            List<double> Value = new List<double>();
            List<string> PatternName = new List<string>();
            List<bool> Replace = new List<bool>();
            int ItemType = 0;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetDataList(In_LoadPat, LoadPat)) { return; }
            if (!DA.GetDataList(In_MyType, MyType)) { return; }
            if (!DA.GetDataList(In_Value, Value)) { return; }
            DA.GetDataList(In_PatternName, PatternName);
            DA.GetDataList(In_Replace, Replace);
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

                if (LoadPat.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        LoadPat.Add(LoadPat[0]);
                    }
                }

                if (MyType.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        MyType.Add(MyType[0]);
                    }
                }

                if (Value.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Value.Add(Value[0]);
                    }
                }
       
                if (PatternName.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        PatternName.Add(PatternName[0]);
                    }
                }

                if (Replace.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Replace.Add(Replace[0]);
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
                    ret = mySapModel.FrameObj.SetLoadTemperature(Name[i], LoadPat[i], MyType[i], Value[i], PatternName[i], Replace[i], itemType);

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AFrameLoadTemp; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("475dba1c-1da2-4997-9f67-bd3ac4cf1135");
    }
}