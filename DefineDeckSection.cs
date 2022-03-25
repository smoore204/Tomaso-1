using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class DefineDeckSection: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public DefineDeckSection()
          : base("Define > Deck > Deck Section", "DfnDckSctn",
            "ETABSv1.SapModel.PropArea.SetDeck()\nInitializes a deck property",
            "Tomaso", "04_Define")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary; }
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_Name;
        int In_DeckType;
        int In_MatProp;
        int In_Thickness;
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
            In_Name = pManager.AddTextParameter("Name", "N", "The name of a deck property. If this is an existing property, that property is modified; otherwise, a new property is added.", 
                GH_ParamAccess.list);
            In_DeckType = pManager.AddIntegerParameter("DeckType", "D", "This is one of the following items in the eDeckType enumeration. \n1. Filled \n2. Unfilled \n3. SolidSlab \nIf this item is Filled, use the Assign Shell Deck Component to get additional parameters.\nIf this item is Unfilled, use the Assign Deck Unfilled component to get additional parameters.\nIf this item is SolidSlab, use the Assign Deck SolidSlab component to get additional parameters.",
                GH_ParamAccess.list);
            In_MatProp = pManager.AddTextParameter("MatProp", "M", "The name of the material property to be assigned to the deck property.",
                GH_ParamAccess.list);
            In_Thickness = pManager.AddNumberParameter("Thickness", "T", "The membrane thickness. [L]",
                GH_ParamAccess.list);
            In_Color = pManager.AddIntegerParameter("Color", "C", "The display color assigned to the property.",
                GH_ParamAccess.list, -1);
            In_Notes = pManager.AddTextParameter("Notes", "N", "The notes, if any, assigned to the property.",
                GH_ParamAccess.list);
            In_GUID = pManager.AddTextParameter("GUID", "G", "The GUID (global unique identifier), if any, assigned to the property.",
                GH_ParamAccess.list);

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
            List<int> DeckType = new List<int>();
            List<string> MatProp = new List<string>();
            List<double> Thickness = new List<double>();
            List<int> Color  = new List<int>();
            List<string> Notes  = new List<string>();
            List<string> GUID = new List<string>();

            if (!DA.GetData(In_Run, ref Run)){return;}
            if (!DA.GetDataList(In_Name, Name)){ return; }
            if (!DA.GetDataList(In_DeckType, DeckType)) { return; }
            if (!DA.GetDataList(In_MatProp, MatProp)) { return; }
            if (!DA.GetDataList(In_Thickness, Thickness)) { return; }
            DA.GetDataList(In_Color, Color );
            DA.GetDataList(In_Notes, Notes );
            DA.GetDataList(In_GUID, GUID);

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

                if (MatProp.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        MatProp.Add(MatProp[0]);
                    }
                }

                if (Thickness.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Thickness.Add(Thickness[0]);
                    }
                }

                if (DeckType.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        DeckType.Add(DeckType[0]);
                    }
                }


                if (Color.Count == 1)
                {
                    for (int i = 0; i < Name.Count - 1; i++)
                    {
                        Color.Add(Color [0]);
                    }
                }

                if (Notes.Count == 0)
                {
                    for (int i = 0; i < Name.Count; i++)
                    {
                        Notes.Add("");
                    }
                }

                if (Notes.Count == 1)
                {
                    for (int i = 0; i < Name.Count; i++)
                    {
                        Notes.Add(Notes[0]);
                    }
                }

                if (GUID.Count == 0)
                {
                    for (int i = 0; i < Name.Count; i++)
                    {
                        GUID.Add("");
                    }
                }

                eDeckType deckType = eDeckType.Filled;


                for (int i = 0; i < Name.Count; i++)
                {
                    switch (DeckType[i])
                    {
                        case 1:
                            deckType = eDeckType.Filled;
                            break;
                        case 2:
                            deckType = eDeckType.Unfilled;
                            break;
                        case 3:
                            deckType = eDeckType.SolidSlab;
                            break;
                    }

                    ret = mySapModel.PropArea.SetDeck(Name[i], deckType, eShellType.Membrane, MatProp[i], Thickness[i], Color[i], Notes[i], GUID[i]);
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.DDeck; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("57be020e-9c57-483e-9800-d9833cf76a65");
    }
}