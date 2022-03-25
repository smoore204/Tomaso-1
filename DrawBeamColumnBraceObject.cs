using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DrawBeamColumnBraceObject: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DrawBeamColumnBraceObject()
          : base("Draw > Beam/Column/Brace Object", "DrwBmClmnBrc",
            "ETABSv1.SapModel.FrameObj.AddByCoord()\nAdds a new frame object whose end points are at the specified coordinates.",
            "Tomaso", "05_Draw")
        {
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Line;
        int In_Name;
        int In_PropName;
        int In_UserName;
        int In_CSys;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Line = pManager.AddLineParameter("Line", "L", "", 
                GH_ParamAccess.list);
            In_Name = pManager.AddTextParameter("Name", "N", "This is the name that the program ultimately assigns for the frame object. If no userName is specified, the program assigns a default name to the frame object. If a userName is specified and that name is not used for another frame, cable or tendon object, the userName is assigned to the frame object, otherwise a default name is assigned to the frame object.", 
                GH_ParamAccess.list);
            In_PropName = pManager.AddTextParameter("PropName", "P", "This is Default, None, or the name of a defined frame section property. If it is Default, the program assigns a default section property to the frame object.If it is None, no section property is assigned to the frame object.If it is the name of a defined frame section property, that property is assigned to the frame object.", 
                GH_ParamAccess.list, "Default");
            In_UserName = pManager.AddTextParameter("UserName", "U", "", 
                GH_ParamAccess.list);
            In_CSys = pManager.AddTextParameter("CSys", "C", "The name of the coordinate system in which the frame object end point coordinates are defined.", 
                GH_ParamAccess.item, "Global");
            
            pManager[In_PropName].Optional = true;
            pManager[In_UserName].Optional = true;
            pManager[In_CSys].Optional = true;
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
            List<Line> Line = new List<Line>();
            List<string> Name = new List<string>();
            List<string> PropName = new List<string>();
            List<string> UserName = new List<string>();
            string CSys = string.Empty;


            if (!DA.GetData(In_Run, ref Run)) { return; }
            if (!DA.GetDataList(In_Line, Line)) { return; }
            if (!DA.GetDataList(In_Name, Name)) { return; }
            DA.GetDataList(In_PropName, PropName);
            DA.GetDataList(In_UserName, UserName);
            DA.GetData(In_CSys, ref CSys);


            if (Line.Count != Name.Count)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Line.Count != Name.Count");
            }


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

                Point3d startPoint;
                Point3d endPoint;

                if (PropName.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        PropName.Add(PropName[0]);
                    }
                }

                if (UserName.Count == 0)
                {
                    for (int i = 0; i < Name.Count; i++)
                    {
                        UserName.Add("");
                    }
                }

                for (int i = 0; i < Line.Count; i++)
                {
                    startPoint = new Point3d(Line[i].To);
                    endPoint = new Point3d(Line[i].From);

                    double XI = startPoint[0];
                    double YI = startPoint[1];
                    double ZI = startPoint[2];
                    double XJ = endPoint[0];
                    double YJ = endPoint[1];
                    double ZJ = endPoint[2];

                    string name = Name[i];

                    ret = mySapModel.FrameObj.AddByCoord(XI, YI, ZI, XJ, YJ, ZJ, ref name, PropName[i], UserName[i], CSys);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.DrawFrame; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("698988F9-03FB-487A-A33F-E95A125E00AC");
    }
}