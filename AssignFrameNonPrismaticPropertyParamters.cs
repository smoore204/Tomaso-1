using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignFrameNonPrismaticPropertyParamters: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignFrameNonPrismaticPropertyParamters()
          : base("Assign > Frame > NonPrismatic Property Parameters", "AssgnFrmNnPrsmtcPrprtyPrmtrs",
            "ETABSv1.SapModel.FrameObj.SetNonPrismatic()\nAssigns data to a nonprismatic frame section property.",
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
        int In_NumberItems;
        int In_Type;
        int In_StartSec;
        int In_EndSec;
        int In_Length;
        int In_EI33;
        int In_EI22;
        int In_Color;
        int In_Notes;
        int In_GUID;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing or new frame section property. If this is an existing property, that property is modified; otherwise, a new property is added.", 
                GH_ParamAccess.list);
            In_NumberItems = pManager.AddIntegerParameter("NumberItems", "N", "The number of segments assigned to the nonprismatic section.",
                GH_ParamAccess.item);
            In_StartSec = pManager.AddTextParameter("StartSec", "S", "This is a list of the names of the frame section properties at the start of each segment. \nAuto select lists and nonprismatic sections are not allowed in this array.",
                GH_ParamAccess.list);
            In_EndSec = pManager.AddTextParameter("EndSec", "E", "This is a list of the names of the frame section properties at the end of each segment. \nAuto select lists and nonprismatic sections are not allowed in this array.",
                GH_ParamAccess.list);
            In_Length = pManager.AddNumberParameter("Length", "L", "This is a list that includes the length of each segment. The length may be variable or absolute as indicated by the MyType item. [L] when length is absolute.",
                GH_ParamAccess.list, 1);
            In_Type = pManager.AddIntegerParameter("Type", "T", "This is a list of either 1 or 2, indicating the length type for each segment:  \n1. Variable (relative length) \n2. Absolute",
                GH_ParamAccess.list);
            In_EI33 = pManager.AddIntegerParameter("EI33", "E3", "This is a list of 1, 2 or 3, indicating the variation type for EI33 in each segment: \n1. Linear \n2. Parbolic \n3. Cubic",
                GH_ParamAccess.list);
            In_EI22 = pManager.AddIntegerParameter("EI22", "E2", "This is a of 1, 2 or 3, indicating the variation type for EI22 in each segment: \n1. Linear \n2. Parbolic \n3. Cubic",
                GH_ParamAccess.list);
            In_Color = pManager.AddIntegerParameter("Color", "C", "The name of the coordinate system for the considered point force load. This is Local or the name of a defined coordinate system.",
                GH_ParamAccess.item, -1);
            In_Notes = pManager.AddTextParameter("Notes", "N", "",
                GH_ParamAccess.item);
            In_GUID = pManager.AddTextParameter("GUID ", "G", "The GUID (global unique identifier), if any, assigned to the section. If this item is input as Default, the program assigns a GUID to the section.",
                GH_ParamAccess.item);

            pManager[In_Color].Optional = true;
            pManager[In_Notes].Optional = true;
            pManager[In_GUID].Optional = true;
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
            int NumberItems = -1;
            List<int> Type = new List<int>();
            List<string> StartSec = new List<string>();
            List<string> EndSec = new List<string>();
            List<double> Length = new List<double>();
            List<int> EI33 = new List<int>();
            List<int> EI22 = new List<int>();
            int Color = -1;
            string Notes  = string.Empty;
            string GUID  = string.Empty;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetData(In_NumberItems, ref NumberItems)) { return; }
            if (!DA.GetDataList(In_Type, Type)) { return; }
            if (!DA.GetDataList(In_StartSec, StartSec)) { return; }
            DA.GetDataList(In_EndSec, EndSec);
            DA.GetDataList(In_Length, Length);
            if (!DA.GetDataList(In_EI33, EI33)) { return; }
            if (!DA.GetDataList(In_EI22, EI22)) { return; }
            DA.GetData(In_Color, ref Color );
            DA.GetData(In_Notes, ref Notes );
            DA.GetData(In_GUID, ref GUID );

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

                string[] startSec = StartSec.ToArray();
                string[] endSec = EndSec.ToArray();
                double[] length = Length.ToArray();
                int[] type = Type.ToArray();
                int[] ei33 = EI33.ToArray();
                int[] ei22 = EI22.ToArray();

                for (int i = 0; i < Name.Count; i++)
                {
                    ret = mySapModel.PropFrame.SetNonPrismatic(Name[i], NumberItems, ref startSec, ref endSec, ref length, ref type, ref ei33, ref ei22, Color, Notes, GUID);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AFrameNonPrismatic; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("b060b867-b072-4576-b326-2b4733b653e0");
    }
}