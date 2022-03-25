using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignFrameOutputStations: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignFrameOutputStations()
          : base("Assign > Frame > Output Stations", "AssgnFrmOtptSttns",
            "ETABSv1.SapModel.FrameObj.SetOutputStations()\nAssigns frame object output station data. ",
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
        int In_MyType;
        int In_MaxSegSize;
        int In_MinSections;
        int In_NoOutPutAndDesignAtElementEnds;
        int In_NoOutPutAndDesignAtPointLoads;
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
            In_MyType = pManager.AddIntegerParameter("Type", "T", "This is 1 or 2, indicating how the output stations are specified. \n1. Maximum segment size, that is, maximum station spacing\n2. Minimum number of stations",
                GH_ParamAccess.list);
            In_MaxSegSize = pManager.AddNumberParameter("MaxSegSize", "M", "The maximum segment size, that is, the maximum station spacing. This item applies only when Type = 1. [L]",
                GH_ParamAccess.list);
            In_MinSections = pManager.AddIntegerParameter("MinSections", "M", "The minimum number of stations. This item applies only when Type = 2.",
                GH_ParamAccess.list);
            In_NoOutPutAndDesignAtElementEnds = pManager.AddBooleanParameter("NoOutPutAndDesignAtElementEnds", "N", "If this item is True, no additional output stations are added at the ends of line elements when the frame object is internally meshed. In ETABS, this item will always be False",
                GH_ParamAccess.list, false);
            In_NoOutPutAndDesignAtPointLoads = pManager.AddBooleanParameter("NoOutPutAndDesignAtPointLoads", "N", "If this item is True, no additional output stations are added at point load locations. In ETABS, this item will always be False.",
                GH_ParamAccess.list, false);
            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0 \nGroup = 1 \nSelectedObjects = 2 \nIf this item is Objects, the restraint assignment is made to the point object specified by the Name item. \nIf this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item. \nIf this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.",
                GH_ParamAccess.item, 0);

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
            List<int> MyType = new List<int>();
            List<double> MaxSegSize = new List<double>();
            List<int> MinSection = new List<int>();
            List<bool> NoOutPutAndDesignAtElementEnds = new List<bool>();
            List<bool> NoOutPutAndDesignAtPointLoads = new List<bool>();
            int ItemType = 0;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetDataList(In_MyType, MyType)) { return; }
            if (!DA.GetDataList(In_MaxSegSize, MaxSegSize)) { return; }
            if (!DA.GetDataList(In_MinSections, MinSection)) { return; }
            if (!DA.GetDataList(In_NoOutPutAndDesignAtElementEnds, NoOutPutAndDesignAtElementEnds)) { return; }
            if (!DA.GetDataList(In_NoOutPutAndDesignAtPointLoads, NoOutPutAndDesignAtPointLoads)) { return; }
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

                if (MyType.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        MyType.Add(MyType[0]);
                    }
                }

                if (MaxSegSize.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        MaxSegSize.Add(MaxSegSize[0]);
                    }
                }

                if (MinSection.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        MinSection.Add(MinSection[0]);
                    }
                }

                if (NoOutPutAndDesignAtElementEnds.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        NoOutPutAndDesignAtElementEnds.Add(NoOutPutAndDesignAtElementEnds[0]);
                    }
                }

                if (NoOutPutAndDesignAtPointLoads.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        NoOutPutAndDesignAtPointLoads.Add(NoOutPutAndDesignAtPointLoads[0]);
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
                    ret = mySapModel.FrameObj.SetOutputStations(Name[i], MyType[i], MaxSegSize[i], MinSection[i], NoOutPutAndDesignAtElementEnds[i], NoOutPutAndDesignAtPointLoads[i], itemType);

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AFrameOutputStation; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("dda83476-21b9-4a13-aa4f-d776e705fa92");
    }
}