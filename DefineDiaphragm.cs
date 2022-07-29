using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DefineDiaphragm: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DefineDiaphragm()
          : base("Define > Diaphragm", "DfnDphrgm",
            "ETABSv1.SapModel.Diaphragm.SetDiaphragm()\nAdds a new diaphragm, or modifies an existing diaphragm ",
            "Tomaso", "04_Define")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.primary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_SemiRigid;

        int Out_Ret;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "This is the name of a diaphragm. If this is the name of an existing diaphragm, that diaphragm is modified, otherwise a new diaphragm is added.",
                GH_ParamAccess.list);
            In_SemiRigid = pManager.AddBooleanParameter("SemiRigid", "S", "This item is True if the diaphragm is to be semi-rigid, and False otherwise. Default is False if no input added.",
                GH_ParamAccess.list);

            pManager[In_SemiRigid].Optional = true;
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
            List<string> Name = new List<string>();
            List<bool> SemiRigid = new List<bool>();

            DA.GetData(In_Run, ref Run);
            if (!DA.GetDataList(In_Name, Name)) { return; }
            DA.GetDataList(In_SemiRigid, SemiRigid);

            if (!SemiRigid.Contains(true) & !SemiRigid.Contains(false))
            {
                for (int i = 0; i < Name.Count; i++)
                    SemiRigid.Add(false);
            }

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
                /////////////////////////////////////////////////

                for (int i = 0; i<Name.Count; i++)
                {
                    ret = mySapModel.Diaphragm.SetDiaphragm(Name[i], SemiRigid[i]);
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
            DA.SetData(Out_Ret, Run);
        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// You can add image files to your project resources and access them like this:
        /// return Resources.IconForThisComponent;
        /// </summary>

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.DDiaphragm; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("a4eec5ba-0418-4e81-97a5-3b7700571d24");
    }
}