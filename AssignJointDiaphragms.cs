using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignJointDiaphragms: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignJointDiaphragms()
          : base("Assign > Joint > Diaphragms", "AssgnJntDphrgms",
            "ETABSv1.SapModel.PointObj.SetDiaphram()\nAssigns a diaphragm to a point object.",
            "Tomaso", "06_Assign")
        {
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_DiaphragmOption;
        int In_DiaphragmName;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing point object.",
                GH_ParamAccess.list);
            In_DiaphragmOption = pManager.AddTextParameter("DiaphragmOption", "D", "This is an item from the eDiaphragmOption enumeration \nIf this item is 'Disconnect' then the point object will be disconnected from any diaphragm \nIf this item is 'FromShellObject' then the point object will inherit the diaphragm assignment of its bounding area object. \nIf this item is 'DefinedDiaphragm' then the point object will be assigned the existing diaphragm specified by DiaphragmName", 
                GH_ParamAccess.list);
            In_DiaphragmName = pManager.AddTextParameter("DiaphragmName", "N", "The name of an existing diaphragm ",
                GH_ParamAccess.list);
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
            List<string> DiaphragmOption = new List<string>();
            List<string> DiaphragmName = new List<string>();


            if (!DA.GetData(In_Run, ref Run)) { return; }
            if (!DA.GetDataList(In_Name, Name)) { return; }
            if (!DA.GetDataList(In_DiaphragmOption, DiaphragmOption)){ return; }
            DA.GetDataList(In_DiaphragmName, DiaphragmName);

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
                
                if (DiaphragmOption.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        DiaphragmOption.Add(DiaphragmOption[0]);
                    }
                }

                if (DiaphragmName.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        DiaphragmName.Add(DiaphragmName[0]);
                    }
                }

                for (int i = 0; i < Name.Count; i++)
                {
                    eDiaphragmOption diaphragmOption = eDiaphragmOption.Disconnect;
                    switch (DiaphragmOption[i])
                    {
                        case "Disconnect":
                            diaphragmOption = eDiaphragmOption.Disconnect;
                            break;
                        case "FromShellObject":
                            diaphragmOption = eDiaphragmOption.FromShellObject;
                            break;
                        case "DefinedDiaphragm":
                            diaphragmOption = eDiaphragmOption.DefinedDiaphragm;
                            break;
                    }
                    if (DiaphragmOption[i] == "DefinedDiaphragm")
                    {
                        ret = mySapModel.PointObj.SetDiaphragm(Name[i], diaphragmOption, DiaphragmName[i]);
                    }
                    else
                    {
                        ret = mySapModel.PointObj.SetDiaphragm(Name[i], diaphragmOption);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AJointDiaphragm; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("788b7405-9191-4e28-a451-d1b4c20ba2ed");
    }
}