using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DrawLinks: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DrawLinks()
          : base("Draw > Links", "DrwLnks",
            "ETABSv1.SapModel.LinkObj.AddByPoint()\nAdds a new link object whose end points are at the specified coordinates. ",
            "Tomaso", "05_Draw")
        {
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Line;
        int In_Point;
        int In_IsSingleJoint;
        int In_Name;
        int In_PropName;
        int In_UserName;
        int In_CSys;

        int Out_Ret;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Line = pManager.AddLineParameter("Line", "L", "Two-joint Link", 
                GH_ParamAccess.list);
            In_Point= pManager.AddTextParameter("PointName", "P", "The name of a defined point object at the end of the added link object for a one-Joint link object",
                GH_ParamAccess.list);
            In_IsSingleJoint = pManager.AddBooleanParameter("IsSingleJoint", "S", "This item is True if a one-joint link is added and False if a two-joint link is added.",
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "This is the name that the program ultimately assigns for the link object. If no UserName is specified, the program assigns a default name to the link object. If a UserName is specified and that name is not used for another link object, the UserName is assigned to the link object; otherwise a default name is assigned to the link object. ", 
                GH_ParamAccess.list);
            In_UserName = pManager.AddTextParameter("UserName", "U", "",
                GH_ParamAccess.list);
            In_PropName = pManager.AddTextParameter("PropName", "P", "This is either Default or the name of a defined link property. If it is Default the program assigns a default link property to the link object.If it is the name of a defined link property, that property is assigned to the link object", 
                GH_ParamAccess.item, "Default");
            In_CSys = pManager.AddTextParameter("CSys", "C", "The name of the coordinate system in which the frame object end point coordinates are defined. ", 
                GH_ParamAccess.item, "Global");

            pManager[In_Line].Optional = true;
            pManager[In_Point].Optional = true;
            pManager[In_IsSingleJoint].Optional = true;
            pManager[In_PropName].Optional = true;
            pManager[In_UserName].Optional = true;
            pManager[In_CSys].Optional = true;
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            Out_Ret = pManager.AddIntegerParameter("Ret", "R", "",
                GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool Run = false;
            List<Line> Line = new List<Line>();
            List<string> Point = new List<string>();
            List<string> Name = new List<string>();
            bool IsSingleJoint = false;
            string PropName = "Default";
            List<string> UserName = new List<string>();
            string CSys = "Global";

            if (!DA.GetData(In_Run, ref Run)){return;}
            DA.GetDataList(In_Line, Line);
            DA.GetDataList(In_Point, Point);
            if (!DA.GetDataList(In_Name, Name)){ return; }
            DA.GetData(In_IsSingleJoint, ref IsSingleJoint);
            DA.GetData(In_PropName, ref PropName);
            DA.GetDataList(In_UserName, UserName);
            DA.GetData(In_CSys, ref CSys);

            string[] username = UserName.ToArray(); 


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

                //add frame object by coordinates
                Point3d startPoint;
                Point3d endPoint;

                if (IsSingleJoint)
                {
                    for (int i = 0; i < Point.Count; i++)
                    {
                        string name = Name[i];

                        if (UserName.Count == 0)
                        {
                            ret = mySapModel.LinkObj.AddByPoint(Point[i], Point[i], ref name, IsSingleJoint, PropName, "");
                        }
                        else
                        {
                            ret = mySapModel.LinkObj.AddByPoint(Point[i], Point[i], ref name, IsSingleJoint, PropName, username[i]);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < Line.Count; i++)
                    {
                        startPoint = new Point3d(Line[i].To);
                        endPoint = new Point3d(Line[i].From);

                        string name = Name[i];

                        if (UserName.Count == 0)
                        {
                            ret = mySapModel.LinkObj.AddByCoord(startPoint.X, startPoint.Y, startPoint.Z, endPoint.X, endPoint.Y, endPoint.Z, ref name, IsSingleJoint, PropName, "", CSys);
                        }
                        else
                        {
                            ret = mySapModel.LinkObj.AddByCoord(startPoint.X, startPoint.Y, startPoint.Z, endPoint.X, endPoint.Y, endPoint.Z, ref name, IsSingleJoint, PropName, username[i], CSys);
                        }

                    }
                }
                //////////////////////////////////////////////////

                //Check ret value
                if (ret != 0)
                {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "API script FAILED to complete.");
                }

                //Refresh View
                ret = mySapModel.View.RefreshView();

                //Clean up variables
                mySapModel = null;
                myETABSObject = null;
                
            }
            DA.SetData(Out_Ret, Run);
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// You can add image files to your project resources and access them like this:
        /// return Resources.IconForThisComponent;
        /// </summary>

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.DrawLink; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("25f22c79-e4e3-4bd2-9383-1338ca26f881");
    }
}