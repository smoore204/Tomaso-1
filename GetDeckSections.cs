using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class GetDeckSections: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public GetDeckSections()
          : base("Get > Deck Sections", "GtDckSctns",
            "ETABSv1.SapModel.PropArea.GetDeck()\nRetrieves property data for Filled, Unfilled, and Solid slab deck section.",
            "Tomaso", "04_Define")
        {
        }
        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.tertiary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;

        int Out_DeckType;
        int Out_SlabDepth;
        int Out_RibDepth;
        int Out_RibWidthTop;
        int Out_RibWidthBot;
        int Out_RibSpacing;
        int Out_ShearThickness;
        int Out_UnitWeight;
        int Out_ShearStudDia;
        int Out_ShearStudHt;
        int Out_ShearStudFu;
        int Out_ShellType;
        int Out_MatProp;
        int Out_Thickness;
        int Out_Color;
        int Out_Notes;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            
            In_Run = pManager.AddBooleanParameter("Run", "R", "", 
                GH_ParamAccess.item, false);
            In_Name = pManager.AddTextParameter("Name", "N", "The name of an existing deck property. ", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            Out_DeckType = pManager.AddGenericParameter("DeckType", "D", "This is one of the items in the eDeckType enumeration.", GH_ParamAccess.item);
            Out_SlabDepth = pManager.AddNumberParameter("SlabDepth", "S", "Slab Depth, tc ", GH_ParamAccess.item);
            Out_RibDepth = pManager.AddNumberParameter("RibDepth", "Rd", "Rib Depth, hr ", GH_ParamAccess.item);
            Out_RibWidthTop = pManager.AddNumberParameter("RibWidthTop", "Wt", "Rib Width Top, wrt ", GH_ParamAccess.item);
            Out_RibWidthBot = pManager.AddNumberParameter("RibWidthBot", "Wb", "Rib Width Bottom, wrb ", GH_ParamAccess.item);
            Out_RibSpacing = pManager.AddNumberParameter("RibSpacing", "Rs", "Rib Spacing, sr ", GH_ParamAccess.item);
            Out_ShearThickness = pManager.AddNumberParameter("ShearThickness", "S", "Deck Shear Thickness ", GH_ParamAccess.item);
            Out_UnitWeight = pManager.AddNumberParameter("UnitWeight", "U", "Deck Unit Weight", GH_ParamAccess.item);
            Out_ShearStudDia = pManager.AddNumberParameter("ShearStudDia", "Sd", "Shear Stud Diameter ", GH_ParamAccess.item);
            Out_ShearStudHt = pManager.AddNumberParameter("ShearStudHt", "Sh", "Shear Stud Height, hs", GH_ParamAccess.item);
            Out_ShearStudFu = pManager.AddNumberParameter("ShearStudFu", "Sf", "Shear Stud Tensile Strength, Fu ", GH_ParamAccess.item);
            Out_ShellType = pManager.AddGenericParameter("Shell Type", "St", "This is one of the items in the eShellType enumeration. Please note that for deck properties, this is always Membrane", GH_ParamAccess.item);
            Out_MatProp = pManager.AddTextParameter("MatProp", "M", "The name of the material property for the area property. This item does not apply when ShellType is Layered.", GH_ParamAccess.item);
            Out_Thickness = pManager.AddNumberParameter("Thickness", "T", "The membrane thickness. [L] This item does not apply when ShellType is Layered.", GH_ParamAccess.item);
            Out_Color = pManager.AddIntegerParameter("Color", "C", "The display color assigned to the property.", GH_ParamAccess.item);
            Out_Notes = pManager.AddTextParameter("Notes", "N", "The notes, if any, assigned to the property.", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool Run = false;
            string Name = string.Empty;

            DA.GetData(In_Run, ref Run);
            DA.GetData(In_Name, ref Name);

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

                eDeckType DeckType = eDeckType.Filled;
                double SlabDepth = -1;
                double RibDepth = -1;
                double RibWidthTop = -1;
                double RibWidthBot = -1;
                double RibSpacing = -1;
                double ShearThickness = -1;
                double UnitWeight = -1;
                double ShearStudDia = -1;
                double ShearStudHt = -1;
                double ShearStudFu = -1;
                eShellType ShellType = eShellType.Membrane;
                string MatProp = string.Empty;
                double Thickness = -1;
                int Color = -1;
                string Notes = string.Empty;
                string GUID = string.Empty;


                /////////////////////////////////////////////////

                ret = mySapModel.PropArea.GetDeck(Name, ref DeckType, ref ShellType, ref MatProp, ref Thickness, ref Color, ref Notes, ref GUID);

                if (DeckType == eDeckType.Unfilled)
                {
                    ret = mySapModel.PropArea.GetDeckUnfilled(Name, ref RibDepth, ref RibWidthTop, ref RibWidthBot, ref RibSpacing, ref ShearThickness, ref UnitWeight);

                }
                if (DeckType == eDeckType.Filled)
                {
                    ret = mySapModel.PropArea.GetDeckFilled(Name, ref SlabDepth, ref RibDepth, ref RibWidthTop, ref RibWidthBot, ref RibSpacing, ref ShearThickness, ref UnitWeight, ref ShearStudDia, ref ShearStudHt, ref ShearStudFu);

                }
                if (DeckType == eDeckType.SolidSlab)
                {
                    ret = mySapModel.PropArea.GetDeckSolidSlab(Name, ref SlabDepth, ref ShearStudDia, ref ShearStudHt, ref ShearStudFu);
                }

                DA.SetData(Out_DeckType, DeckType);
                DA.SetData(Out_ShellType, ShellType);
                DA.SetData(Out_MatProp, MatProp);
                DA.SetData(Out_Thickness, Thickness);
                DA.SetData(Out_Color, Color);
                DA.SetData(Out_Notes, Notes);
                DA.SetData(Out_SlabDepth, SlabDepth);
                DA.SetData(Out_RibDepth, RibDepth);
                DA.SetData(Out_RibWidthTop, RibWidthTop);
                DA.SetData(Out_RibWidthBot, RibWidthBot);
                DA.SetData(Out_RibSpacing, RibSpacing);
                DA.SetData(Out_ShearThickness, ShearThickness);
                DA.SetData(Out_UnitWeight, UnitWeight);
                DA.SetData(Out_ShearStudDia, ShearStudDia);
                DA.SetData(Out_ShearStudHt, ShearStudHt);
                DA.SetData(Out_ShearStudFu, ShearStudFu);

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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.SteveMemoji; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("9931abe8-f539-40c1-adae-d3018553a0ea");
    }
}