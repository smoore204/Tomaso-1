using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignFrameSection: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignFrameSection()
          : base("Assign > Frame > Section", "AssgnFrmSctn",
            "ETABSv1.SapModel.FrameObj.SetSection()\nAssigns a frame section property to a frame object.",
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
        int In_PropName;
        int In_ItemType;
        int In_SVarRelStartLoc;
        int In_SVarTotalLength;


        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing frame object or group depending on the value of the ItemType item.", 
                GH_ParamAccess.list);
            In_PropName = pManager.AddTextParameter("PropName", "P", "This is None or the name of a frame section property to be assigned to the specified frame object(s).",
                GH_ParamAccess.list);
            In_SVarRelStartLoc = pManager.AddNumberParameter("SVarRelStartLoc", "S", "",
                GH_ParamAccess.list, 0);
            In_SVarTotalLength = pManager.AddNumberParameter("SVarTotalLength", "S", "",
                GH_ParamAccess.list, 0);
            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0 \nGroup = 1 \nSelectedObjects = 2 \nIf this item is Objects, the restraint assignment is made to the point object specified by the Name item. \nIf this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item. \nIf this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.",
                GH_ParamAccess.item, 0);

            pManager[In_ItemType].Optional = true;
            pManager[In_SVarRelStartLoc].Optional = true;
            pManager[In_SVarTotalLength].Optional = true;
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
            List<string> PropName = new List<string>();
            List<double> SVarRelStartLoc = new List<double>();
            List<double> SVarTotalLength = new List<double>();
            int ItemType = 0;


            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetDataList(In_PropName, PropName)){ return; };
            if (!DA.GetDataList(In_SVarRelStartLoc, SVarRelStartLoc)) { return; };
            if (!DA.GetDataList(In_SVarTotalLength, SVarTotalLength)) { return; };
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

                if (SVarRelStartLoc.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        SVarRelStartLoc.Add(SVarRelStartLoc[0]);
                    }
                }

                if (SVarTotalLength.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        SVarTotalLength.Add(SVarTotalLength[0]);
                    }
                }

                if (PropName.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        PropName.Add(PropName[0]);
                    }
                }

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
                    ret = mySapModel.FrameObj.SetSection(Name[i], PropName[i], itemType, SVarRelStartLoc[i], SVarTotalLength[i]);
                   
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AFrameSection; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("d312865b-cb80-4676-b397-519dde9d863f");
    }
}