using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignJointSprings: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignJointSprings()
          : base("Assign > Joint > Springs", "AssgnJntSprngs",
            "ETABSv1.SapModel.PointObj.SetSpring()\nAssigns coupled springs to a point object.",
            "Tomaso", "06_Assign")
        {
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_K;
        int In_ItemType;
        int In_IsLocalCsys;
        int In_Replace;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing point object or group depending on the value of the ItemType item.", 
                GH_ParamAccess.list);
            In_K = pManager.AddNumberParameter("K", "K", "This is an array of six spring stiffness values. \nValue(0) = U1[F / L]\nValue(1) = U2[F / L]\nValue(2) = U3[F / L]\nValue(3) = R1[FL / rad]\nValue(4) = R2[FL / rad]\nValue(5) = R3[FL / rad]",
                GH_ParamAccess.list);
            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0 \nGroup = 1 \nSelectedObjects = 2 \nIf this item is Objects, the restraint assignment is made to the point object specified by the Name item. \nIf this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item. \nIf this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.",
                GH_ParamAccess.item, 0);
            In_IsLocalCsys = pManager.AddBooleanParameter("IsLocalCsys", "L", "If this item is True, the specified spring assignments are in the point object local coordinate system. If it is False, the assignments are in the Global coordinate s",
                GH_ParamAccess.item, true);
            In_Replace = pManager.AddBooleanParameter("Replace", "R", "If this item is True, all existing point spring assignments to the specified point object(s) are deleted prior to making the assignment. If it is False, the spring assignments are added to any existing assignments.",
                GH_ParamAccess.item, false);

            pManager[In_ItemType].Optional = true;
            pManager[In_IsLocalCsys].Optional = true;
            pManager[In_Replace].Optional = true;
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
            List<double> K = new List<double>();
            int ItemType = 0;
            bool IsLocalCsys = true;
            bool Replace = false;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            DA.GetDataList(In_K, K);
            DA.GetData(In_ItemType, ref ItemType);
            DA.GetData(In_IsLocalCsys, ref IsLocalCsys);
            DA.GetData(In_Replace, ref Replace);


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

                double[] k = K.ToArray();
                string [] NameArr = Name.ToArray();

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

                    ret = mySapModel.PointObj.SetSpring(NameArr[i], ref k, itemType, IsLocalCsys, Replace);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AJointSpring; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("c2018cae-4764-49ca-aaaf-86475689df47");
    }
}