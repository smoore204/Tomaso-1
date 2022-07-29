using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DefineGroupDefinitions: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DefineGroupDefinitions()
          : base("Define > Group Definitions", "DfnGrpDfntns",
            "ETABSv1.SapModel.Group.SetGroup_1()\nSets the group data. Primarily for ETABS.",
            "Tomaso", "04_Define")
        {
        }
        
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.quarternary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_Color;
        int In_SpecifiedForSelection;
        int In_SpecifiedForSectionCutDefinition;
        int In_SpecifiedForSteelDesign;
        int In_SpecifiedForConcreteDesign;
        int In_SpecifiedForAluminumDesign;
        int In_SpecifiedForStaticNLActiveStage;
        int In_SpecifiedForAutoSeismicOutput;
        int In_SpecifiedForAutoWindOutput;
        int In_SpecifiedForMassAndWeight;
        int In_SpecifiedForSteelJoistDesign;
        int In_SpecifiedForWallDesign;
        int In_SpecifiedForBasePlateDesign;
        int In_SpecifiedForConnectionDesign;

        int Out_Ret;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing object or group depending on the value of the ItemType item.", 
                GH_ParamAccess.list);
            In_Color = pManager.AddIntegerParameter("Color", "C", "The display color for the group specified as a Integer. If this value is input as –1, the program automatically selects a display color for the group.",
                GH_ParamAccess.item, -1);
            In_SpecifiedForSelection = pManager.AddBooleanParameter("SpecifiedForSelection", "S", "This item is True if the group is specified to be used for selection; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForSectionCutDefinition = pManager.AddBooleanParameter("SpecifiedForSectionCutDefinition", "S", "This item is True if the group is specified to be used for defining section cuts; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForSteelDesign = pManager.AddBooleanParameter("SpecifiedForSteelDesign", "S", "This item is True if the group is specified to be used for defining steel frame design groups; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForConcreteDesign = pManager.AddBooleanParameter("SpecifiedForConcreteDesign", "C", "This item is True if the group is specified to be used for defining concrete frame design groups; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForAluminumDesign = pManager.AddBooleanParameter("SpecifiedForAluminumDesig", "A", "This item is True if the group is specified to be used for defining aluminum frame design groups; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForStaticNLActiveStage = pManager.AddBooleanParameter("SpecifiedForStaticNLActiveStag", "S", "This item is True if the group is specified to be used for defining stages for nonlinear static analysis; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForAutoSeismicOutput = pManager.AddBooleanParameter("SpecifiedForAutoSeismicOutput", "A", "This item is True if the group is specified to be used for reporting auto seismic loads; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForAutoWindOutput = pManager.AddBooleanParameter("SpecifiedForAutoWindOutput", "A", "This item is True if the group is specified to be used for reporting auto wind loads; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForMassAndWeight = pManager.AddBooleanParameter("SpecifiedForMassAndWeight", "M", "This item is True if the group is specified to be used for reporting group masses and weight; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForSteelJoistDesign = pManager.AddBooleanParameter("SpecifiedForSteelJoistDesign", "S", "This item is True if the group is specified to be used for defining steel joist design groups; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForWallDesign = pManager.AddBooleanParameter("SpecifiedForWallDesign", "W", "This item is True if the group is specified to be used for defining wall design groups; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForBasePlateDesign = pManager.AddBooleanParameter("SpecifiedForBasePlateDesig", "B", "This item is True if the group is specified to be used for defining base plate design groups; otherwise it is False.",
                GH_ParamAccess.item, false);
            In_SpecifiedForConnectionDesign = pManager.AddBooleanParameter("SpecifiedForConnectionDesign", "C", "This item is True if the group is specified to be used for defining connection design groups; otherwise it is False.",
                GH_ParamAccess.item, false);

            pManager[In_Color].Optional = true;
            pManager[In_SpecifiedForSelection].Optional = true;
            pManager[In_SpecifiedForSectionCutDefinition].Optional = true;
            pManager[In_SpecifiedForSteelDesign].Optional = true;
            pManager[In_SpecifiedForConcreteDesign].Optional = true;
            pManager[In_SpecifiedForAluminumDesign].Optional = true;
            pManager[In_SpecifiedForStaticNLActiveStage].Optional = true;
            pManager[In_SpecifiedForAutoSeismicOutput].Optional = true;
            pManager[In_SpecifiedForAutoWindOutput].Optional = true;
            pManager[In_SpecifiedForMassAndWeight].Optional = true;
            pManager[In_SpecifiedForSteelJoistDesign].Optional = true;
            pManager[In_SpecifiedForWallDesign].Optional = true;
            pManager[In_SpecifiedForBasePlateDesign].Optional = true;
            pManager[In_SpecifiedForConnectionDesign].Optional = true;
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
            int Color = -1;
            bool SpecifiedForSelection = false;
            bool SpecifiedForSectionCutDefinition = false;
            bool SpecifiedForSteelDesign = false;
            bool SpecifiedForConcreteDesign = false;
            bool SpecifiedForAluminumDesign = false;
            bool SpecifiedForStaticNLActiveStage = false;
            bool SpecifiedForAutoSeismicOutput = false;
            bool SpecifiedForAutoWindOutput = false;
            bool SpecifiedForMassAndWeight = false;
            bool SpecifiedForSteelJoistDesign = false;
            bool SpecifiedForWallDesign = false;
            bool SpecifiedForBasePlateDesign = false;
            bool SpecifiedForConnectionDesign = false;

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            DA.GetData(In_Color, ref Color);
            DA.GetData(In_SpecifiedForSelection, ref SpecifiedForSelection);
            DA.GetData(In_SpecifiedForSectionCutDefinition, ref SpecifiedForSectionCutDefinition);
            DA.GetData(In_SpecifiedForSteelDesign, ref SpecifiedForSteelDesign);
            DA.GetData(In_SpecifiedForConcreteDesign, ref SpecifiedForConcreteDesign);
            DA.GetData(In_SpecifiedForAluminumDesign, ref SpecifiedForAluminumDesign);
            DA.GetData(In_SpecifiedForStaticNLActiveStage, ref SpecifiedForStaticNLActiveStage);
            DA.GetData(In_SpecifiedForAutoSeismicOutput, ref SpecifiedForAutoSeismicOutput);
            DA.GetData(In_SpecifiedForAutoWindOutput, ref SpecifiedForAutoWindOutput);
            DA.GetData(In_SpecifiedForMassAndWeight, ref SpecifiedForMassAndWeight);
            DA.GetData(In_SpecifiedForSteelJoistDesign, ref SpecifiedForSteelJoistDesign);
            DA.GetData(In_SpecifiedForWallDesign, ref SpecifiedForWallDesign);
            DA.GetData(In_SpecifiedForBasePlateDesign, ref SpecifiedForBasePlateDesign);
            DA.GetData(In_SpecifiedForConnectionDesign, ref SpecifiedForConnectionDesign);

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

                for (int i = 0; i < Name.Count; i++)
                {
                    ret = mySapModel.GroupDef.SetGroup_1(Name[i], Color, SpecifiedForSelection, SpecifiedForSectionCutDefinition, SpecifiedForSteelDesign, SpecifiedForConcreteDesign, SpecifiedForAluminumDesign, SpecifiedForStaticNLActiveStage, SpecifiedForAutoSeismicOutput, SpecifiedForAutoWindOutput, SpecifiedForMassAndWeight, SpecifiedForSteelJoistDesign, SpecifiedForWallDesign, SpecifiedForBasePlateDesign, SpecifiedForConnectionDesign);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.DGroup; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("a8fd3d73-5266-4f31-81cc-9a775249347c");
    }
}