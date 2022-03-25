using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignFrameLoadsDistributed: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignFrameLoadsDistributed()
          : base("Assign > Frame Loads > Distributed", "AssgnFrmLdsDstrbtd",
            "ETABSv1.SapModel.FrameObj.SetLoadDistributed()\nAssigns distributed loads to frame objects. ",
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
        int In_Dir;
        int In_Dist1;
        int In_Dist2;
        int In_Val1;
        int In_Val2;
        int In_CSys;
        int In_RelDist;
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
            In_MyType = pManager.AddIntegerParameter("Type", "T", "This is 1 or 2, indicating the type of distributed load. \n1. Force per unit length \n1. Moment per unit length",
                GH_ParamAccess.list);
            In_Dir = pManager.AddIntegerParameter("Dir", "D", "This is an integer between 1 and 11, indicating the direction of the load. \n1. Local 1 axis(only applies when CSys is Local) \n2. Local 2 axis(only applies when CSys is Local) \n3. Local 3 axis(only applies when CSys is Local) \n4. X direction(does not apply when CSys is Local) \n5. Y direction(does not apply when CSys is Local) \n6. Z direction(does not apply when CSys is Local) \n7. Projected X direction(does not apply when CSys is Local) \n8. Projected Y direction(does not apply when CSys is Local) \n9. Projected Z direction(does not apply when CSys is Local) \n10. Gravity direction(only applies when CSys is Global) \n11. Projected Gravity direction(only applies when CSys is Global) \nThe positive gravity direction(see Dir = 10 and 11) is in the negative Global Z direction.",
                GH_ParamAccess.list);
            In_Dist1 = pManager.AddNumberParameter("Dist1", "D1", "This is the distance from the I-End of the frame object to the start of the distributed load. This may be a relative distance (0 <= Dist1 <= 1) or an actual distance, depending on the value of the RelDist item. [L] when RelDist is False ",
                GH_ParamAccess.list, 0);
            In_Dist2 = pManager.AddNumberParameter("Dist2", "D2", "This is the distance from the J-End of the frame object to the start of the distributed load. This may be a relative distance (0 <= Dist1 <= 1) or an actual distance, depending on the value of the RelDist item. [L] when RelDist is False ",
                GH_ParamAccess.list, 1);
            In_Val1 = pManager.AddNumberParameter("Val1", "V1", "This is the load value at the start of the distributed load. \n[F/L] when Type is 1 and [M/L] when Type is 2 ",
                GH_ParamAccess.list);
            In_Val2 = pManager.AddNumberParameter("Val2", "V2", "This is the load value at the end of the distributed load. \n[F/L] when Type is 1 and [M/L] when Type is 2 ",
                GH_ParamAccess.list);
            In_CSys = pManager.AddTextParameter("CSys", "C", "The name of the coordinate system for the considered point force load. This is Local or the name of a defined coordinate system.",
                GH_ParamAccess.list, "Global");
            In_RelDist = pManager.AddBooleanParameter("RelDist", "R", "If this item is True, the specified Dist item is a relative distance, otherwise it is an actual distance. ",
                GH_ParamAccess.list, true);
            In_Replace = pManager.AddBooleanParameter("Replace", "R", "If this item is True, all previous loads, if any, assigned to the specified point object(s) in the specified load pattern are deleted before making the new assignment. ",
                GH_ParamAccess.list, true);
            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0 \nGroup = 1 \nSelectedObjects = 2 \nIf this item is Objects, the restraint assignment is made to the point object specified by the Name item. \nIf this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item. \nIf this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.",
                GH_ParamAccess.item, 0);

            pManager[In_Dist1].Optional = true;
            pManager[In_Dist1].Optional = true;
            pManager[In_CSys].Optional = true;
            pManager[In_RelDist].Optional = true;
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
            List<int> Dir = new List<int>();
            List<double> Dist1 = new List<double>();
            List<double> Dist2 = new List<double>();
            List<double> Val1 = new List<double>();
            List<double> Val2 = new List<double>();
            List<string> CSys = new List<string>();
            List<bool> RelDist = new List<bool>();
            List<bool> Replace = new List<bool>();
            int ItemType = 0;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetDataList(In_LoadPat, LoadPat)) { return; }
            if (!DA.GetDataList(In_MyType, MyType)) { return; }
            if (!DA.GetDataList(In_Dir, Dir)) { return; }
            DA.GetDataList(In_Dist1, Dist1);
            DA.GetDataList(In_Dist2, Dist2);
            if (!DA.GetDataList(In_Val1, Val1)) { return; }
            if (!DA.GetDataList(In_Val2, Val2)) { return; }
            DA.GetDataList(In_CSys, CSys);
            DA.GetDataList(In_RelDist, RelDist);
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

                if (Dir.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Dir.Add(Dir[0]);
                    }
                }

                if (Dist1.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Dist1.Add(Dist1[0]);
                    }
                }

                if (Dist2.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Dist2.Add(Dist2[0]);
                    }
                }

                if (Val1.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Val1.Add(Val1[0]);
                    }
                }

                if (Val2.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Val2.Add(Val2[0]);
                    }
                }

                if (CSys.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        CSys.Add(CSys[0]);
                    }
                }
       
                if (RelDist.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        RelDist.Add(RelDist[0]);
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
                    ret = mySapModel.FrameObj.SetLoadDistributed(Name[i], LoadPat[i], MyType[i], Dir[i], Dist1[i], Dist2[i], Val1[i], Val2[i], CSys[i], RelDist[i], Replace[i], itemType);

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AFrameLoadDist; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("c1368bd6-ea57-4759-a13b-4b7153cb560b");
    }
}