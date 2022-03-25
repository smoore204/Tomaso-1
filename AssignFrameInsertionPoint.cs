using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignFrameInsertionPoint: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignFrameInsertionPoint()
          : base("Assign > Frame > Insertion Point", "AssgnFrmInsrtnPt",
            "ETABSv1.SapModel.FrameObj.SetInsertionPoint()\nAssigns frame object insertion point data. The assignments include the cardinal point and end joint offsets.",
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
        int In_CardinalPoint;
        int In_Mirror2;
        int In_StiffTransform;
        int In_Offset1;
        int In_Offset2;
        int In_CSys;
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
            In_CardinalPoint = pManager.AddIntegerParameter("CardinalPoint", "C", "This is a numeric value from 1 to 11 that specifies the cardinal point for the frame object. The cardinal point specifies the relative position of the frame section on the line representing the frame object. \n1. bottom left\n2. bottom center\n3. bottom right\n4. middle left\n5. middle center\n6. middle right\n7. top left\n8. top center\n9. top right\n10.centroid\n11. shear center",
                GH_ParamAccess.item);
            In_Mirror2 = pManager.AddBooleanParameter("Mirror2", "M", "If this item is True, the frame object section is assumed to be mirrored (flipped) about its local 2-axis. ",
                GH_ParamAccess.item);
            In_StiffTransform = pManager.AddBooleanParameter("StiffTransform", "S", "If this item is True, the frame object stiffness is transformed for cardinal point and joint offsets from the frame section centroid. ",
                GH_ParamAccess.item);
            In_Offset1 = pManager.AddNumberParameter("Offset1", "O1", "This is a list of three joint offset distances, in the coordinate directions specified by CSys, at the I-End of the frame object. [L] \nOffset1(0) = Offset in the 1 - axis or X - axis direction\nOffset1(1) = Offset in the 2 - axis or Y - axis direction\nOffset1(2) = Offset in the 3 - axis or Z - axis direction",
                GH_ParamAccess.list);
            In_Offset2 = pManager.AddNumberParameter("Offset2", "O2", "This is a list of three joint offset distances, in the coordinate directions specified by CSys, at the J-End of the frame object. [L] \nOffset1(0) = Offset in the 1 - axis or X - axis direction\nOffset1(1) = Offset in the 2 - axis or Y - axis direction\nOffset1(2) = Offset in the 3 - axis or Z - axis direction",
                GH_ParamAccess.list);
            In_CSys = pManager.AddTextParameter("CSys", "C", "The name of the coordinate system for the considered point force load. This is Local or the name of a defined coordinate system.",
                GH_ParamAccess.item, "Local");
            In_ItemType = pManager.AddIntegerParameter("ItemType", "I", "This is one of the following items in the eItemType enumeration: \nObject = 0 \nGroup = 1 \nSelectedObjects = 2 \nIf this item is Objects, the restraint assignment is made to the point object specified by the Name item. \nIf this item is Group, the restraint assignment is made to all point objects in the group specified by the Name item. \nIf this item is SelectedObjects, the restraint assignment is made to all selected point objects and the Name item is ignored.",
                GH_ParamAccess.item, 0);

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
            int CardinalPoint = -1;
            bool Mirror2 = false;
            bool StiffTransform = false;
            List<double> Offset1 = new List<double>();
            List<double> Offset2 = new List<double>();
            string CSys = string.Empty;
            int ItemType = 0;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetData(In_CardinalPoint, ref CardinalPoint)) { return; }
            if (!DA.GetData(In_Mirror2, ref Mirror2)) { return; }
            if (!DA.GetData(In_StiffTransform, ref StiffTransform)) { return; }
            if (!DA.GetDataList(In_Offset1, Offset1)) { return; }
            if (!DA.GetDataList(In_Offset2, Offset2)) { return; }
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
                double[] offset1 = Offset1.ToArray();
                double[] offset2 = Offset1.ToArray();

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
                    ret = mySapModel.FrameObj.SetInsertionPoint(Name[i], CardinalPoint, Mirror2, StiffTransform, ref offset1, ref offset2, CSys, itemType);

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AFrameInsertion; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("b30d57b5-059c-4b6a-ae06-3d1ec2e29cb3");
    }
}