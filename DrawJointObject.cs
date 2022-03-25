using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DrawJointObject: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DrawJointObject()
          : base("Draw > Joint Objects", "DrwJntObjcts",
            "ETABSv1.SapModel.PointObj.AddCartesian()\nAdds a point object to a model. The added point object will be tagged as a Special Point except if it was merged with another point object. Special points are allowed to exist in the model with no objects connected to them.",
            "Tomaso", "05_Draw")
        {
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Point;
        int In_Name;
        int In_UserName;
        int In_CSys;
        int In_MergeOff;
        int In_MergeNumber;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Point = pManager.AddPointParameter("Point", "P", "", 
                GH_ParamAccess.list);
            In_Name = pManager.AddTextParameter("Name", "N", "This is the name that the program ultimately assigns for the joint object. If no userName is specified, the program assigns a default name to the joint object. If a userName is specified and that name is not used for another joint object, the userName is assigned to the joint object; otherwise a default name is assigned to the joint object.", 
                GH_ParamAccess.list);
            In_UserName = pManager.AddTextParameter("UserName", "U", "This is an optional user specified name for the point object. If a UserName is specified and that name is already used for another point object, the program ignores the UserName.", 
                GH_ParamAccess.list);
            In_CSys = pManager.AddTextParameter("CSys", "C", "The name of the coordinate system in which the frame object end point coordinates are defined.", 
                GH_ParamAccess.item, "Global");
            In_MergeOff = pManager.AddBooleanParameter("MergeOff", "M", "If this item is False, a new point object that is added at the same location as an existing point object will be merged with the existing point object (assuming the two point objects have the same MergeNumber) and thus only one point object will exist at the location. If this item is True, the points will not merge and two point objects will exist at the same location.",
                GH_ParamAccess.item, false);
            In_MergeNumber = pManager.AddIntegerParameter("MergeNumber", "M", "Two points objects in the same location will merge only if their merge number assignments are the same. By default all point objects have a merge number of zero.",
                GH_ParamAccess.item, 0);

            pManager[In_UserName].Optional = true;
            pManager[In_CSys].Optional = true;
            pManager[In_MergeOff].Optional = true;
            pManager[In_MergeNumber].Optional = true;
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
            List<Point3d> Point = new List<Point3d>();
            List<string> Name = new List<string>();
            List<string> UserName = new List<string>();
            string CSys = "Global";
            bool MergeOff = false;
            int MergeNumber = 0;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Point, Point)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            DA.GetDataList(In_UserName, UserName);
            DA.GetData(In_CSys, ref CSys);
            DA.GetData(In_MergeOff, ref MergeOff);
            DA.GetData(In_MergeNumber, ref MergeNumber);

            string[] nameArr = Name.ToArray();
            string[] usernameArr = UserName.ToArray();

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

                //Add frame object by coordinates
                for (int i = 0; i < Point.Count; i++) 
                {
                    string name = nameArr[i];

                    if (UserName.Count == 0)
                    {
                        ret = mySapModel.PointObj.AddCartesian(Point[i][0], Point[i][1], Point[i][2], ref name, "", CSys, MergeOff, MergeNumber);
                    }
                    else
                    {
                        ret = mySapModel.PointObj.AddCartesian(Point[i][0], Point[i][1], Point[i][2], ref name, UserName[i], CSys, MergeOff, MergeNumber);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.DrawPoint; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("dca0f75b-1d3f-4b6f-b7c5-37150bd4e8bb");
    }
}