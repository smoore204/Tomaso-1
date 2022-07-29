using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class AssignShellWallAutoMeshOptions: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignShellWallAutoMeshOptions()
          : base("Assign > Shell > Wall Auto Mesh Options", "AssnShllAtWllMshOptns",
            "ETABSv1.SapModel.DatabaseTables.SetTableForEditingArray()\nAssign Shell Wall Auto Mesh Options",
            "Tomaso", "06_Assign")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.tertiary; }
        }

        //Initialize In and Out Parameters

        int In_Run;
        int In_MeshOption;
        int In_Nv;
        int In_Nh;
        int In_AddRestraints;
        int In_Name;


        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "",
                GH_ParamAccess.list);
            In_MeshOption = pManager.AddTextParameter("MeshOption", "M", "'Default', 'Auto Rectangular', or 'Mesh Nv x Nh'", 
                GH_ParamAccess.item);
            In_AddRestraints = pManager.AddTextParameter("AddRestraints", "A", "'Yes' or 'No'", 
                GH_ParamAccess.item);
            In_Nv = pManager.AddTextParameter("Nv", "V", "Number of vertical sections, only required if MeshOption = Mesh Nv x Nh",
                GH_ParamAccess.item);
            In_Nh = pManager.AddTextParameter("Nh", "H", "Number of horizontal sections, only required if MeshOption = Mesh Nv x Nh",
                GH_ParamAccess.item);

            pManager[In_Nh].Optional = true;
            pManager[In_Nv].Optional = true;

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
            string MeshOption = string.Empty;
            string AddRestraints = string.Empty;
            string Nv = string.Empty;
            string Nh = string.Empty;

            if (!DA.GetData(In_Run, ref Run)) { return; }
            if (!DA.GetDataList(In_Name, Name)) { return; }
            if (!DA.GetData(In_MeshOption, ref MeshOption)) { return; }
            if (!DA.GetData(In_AddRestraints, ref AddRestraints)) { return; }
            if (!DA.GetData(In_Nv, ref Nv)) { return; }
            if (!DA.GetData(In_Nh, ref Nh)) { return; }

            //Test matching number of inputs

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
                ///
                int TableVersion = 0;
                string TableKey = "Area Assignments - Wall Auto Mesh Options";
                string[] FieldsKeysIncluded = { "UniqueName", "Mesh Option", "Nv", "Nh", "Add Restraints" };
                int NumberRecords = Name.Count*5;
                List<string> Data = new List<string>();

                for (int i = 0; i < Name.Count; i++)
                {
                    if (MeshOption == "Default")
                    {
                        List<string> data = new List<string> { Name[i], MeshOption, "", "", AddRestraints };
                        Data.AddRange(data);
                    }
                    else if (MeshOption == "Mesh Nv x Nh")
                    {
                        List<string> data = new List<string> { Name[i], MeshOption, Nv, Nh, AddRestraints };
                        Data.AddRange(data);
                    }
                    else if (MeshOption == "Auto Rectangular")
                    {
                        List<string> data = new List<string> { Name[i], MeshOption, "", "", AddRestraints };
                        Data.AddRange(data);
                    }
                }

                string[] DataArr = Data.ToArray();
                ret = mySapModel.DatabaseTables.SetTableForEditingArray(TableKey, ref TableVersion, ref FieldsKeysIncluded, NumberRecords, ref DataArr);

                bool FillImport = true;
                int NumFatalErrors = 0;
                int NumErrorMsgs = 0;
                int NumWarnMsgs = 0;
                int NumInfoMsgs = 0;
                string ImportLog = "";

                ret = mySapModel.DatabaseTables.ApplyEditedTables(FillImport, ref NumFatalErrors, ref NumErrorMsgs, ref NumWarnMsgs, ref NumInfoMsgs, ref ImportLog);

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AShellWallMesh ; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("a7c20220-8ef9-4f0b-9654-4fe757b8c5e9");
    }
}