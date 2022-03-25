using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;

namespace Tomaso
{
    public class EditStories: GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public EditStories()
          : base("Edit > Stories", "EditStories",
            "ETABSv1.SapModel.Stories.SetStories_2()\nSets the stories for the current tower.",
            "Tomaso", "02_Edit")
        {
        }

        //Initialize In and Out Parameters
        int In_Run;
        int In_BaseElev;
        int In_NumStories;
        int In_StoryNames;
        int In_StoryHeights;
        int In_IsMasterStory;
        int In_SimilarToStory;
        int In_SpliceAbove;
        int In_SpliceHeight;
        int In_Color;

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            In_Run = pManager.AddBooleanParameter("Run", "R", "",
                GH_ParamAccess.item);
            In_BaseElev = pManager.AddNumberParameter("BaseElevation", "B", "The elevation of the base [L]",
                GH_ParamAccess.item, 0);
            In_NumStories = pManager.AddIntegerParameter("NumberStories", "N", "The number of stories",
                GH_ParamAccess.item);
            In_StoryNames = pManager.AddTextParameter("StoryNames", "SN", "The names of the stories ", 
                GH_ParamAccess.list);
            In_StoryHeights = pManager.AddNumberParameter("StoryHeights", "SH", "The story heights [L] ", 
                GH_ParamAccess.list);
            In_IsMasterStory = pManager.AddBooleanParameter("IsMasterStory", "M", "Whether the story is a master story. If no input given, default is set at 'True' for first story and 'False' for all subsequent.", 
                GH_ParamAccess.list);
            In_SimilarToStory = pManager.AddTextParameter("SimilarToStory", "S", "If the story is not a master story, which master story the story is similar to. If no input given, default set to 'None' for all stories.",
                GH_ParamAccess.list, "None");
            In_SpliceAbove = pManager.AddBooleanParameter("SpliceAbove", "SA", "This is True if the story has a splice height, and False otherwise. If no input given, default set to 'False' for all stories.",
                GH_ParamAccess.list, false);
            In_SpliceHeight = pManager.AddNumberParameter("SpliceHeight", "SH", "The story splice height [L]. If no input given, default set to '0' for all stories.",
                GH_ParamAccess.list, 0);
            In_Color = pManager.AddIntegerParameter("Color", "C", "The display color for the story specified as an Integer. If no input given, default color equivalet to '255' (black) for all stories.",
                GH_ParamAccess.list, 255);

            pManager[In_BaseElev].Optional = true;
            pManager[In_IsMasterStory].Optional = true;
            pManager[In_SimilarToStory].Optional = true;
            pManager[In_SpliceAbove].Optional = true;
            pManager[In_SpliceHeight].Optional = true;
            pManager[In_Color].Optional = true;
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
            double BaseElevation = -1;
            int NumberStories = -1;
            List<string> StoryNames = new List<string>();
            List<double> StoryHeights = new List<double>();
            List<bool> IsMasterStory = new List<bool>();
            List<string> SimilarToStory = new List<string>();
            List<bool> SpliceAbove = new List<bool>();
            List<double> SpliceHeight = new List<double>();
            List<int> Color = new List<int>();

            DA.GetData(In_Run, ref Run);
            if (!DA.GetData(In_BaseElev, ref BaseElevation)) { return; }
            if (!DA.GetData(In_NumStories, ref NumberStories)) { return; }
            if (!DA.GetDataList(In_StoryNames, StoryNames)) { return; }
            if (!DA.GetDataList(In_StoryHeights, StoryHeights)) { return; }
            DA.GetDataList(In_IsMasterStory, IsMasterStory);
            DA.GetDataList(In_SimilarToStory, SimilarToStory);
            DA.GetDataList(In_SpliceAbove, SpliceAbove);
            DA.GetDataList(In_SpliceHeight, SpliceHeight);
            DA.GetDataList(In_Color, Color);

            //Initialize default array values

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
                string[] storyNames = StoryNames.ToArray();
                double[] storyHeights = StoryHeights.ToArray();

                if (IsMasterStory.Count == 0)
                {
                    IsMasterStory.Add(true);
                    for (int i = 2; i < StoryNames.Count; i++)
                    {
                        IsMasterStory.Add(false);
                    }
                }

                if (SimilarToStory.Count == 1)
                {
                    for (int i = 1; i < StoryNames.Count; i++)
                    {
                        SimilarToStory.Add(SimilarToStory[0]);
                    }
                }

                if (SpliceAbove.Count == 1)
                {
                    for (int i = 1; i < StoryNames.Count; i++)
                    {
                        SpliceAbove.Add(SpliceAbove[0]);
                    }
                }

                if (SpliceHeight.Count == 1)
                {
                    for (int i = 1; i < StoryNames.Count; i++)
                    {
                        SpliceHeight.Add(SpliceHeight[0]);
                    }
                }

                if (Color.Count == 1)
                {
                    for (int i = 1; i < StoryNames.Count; i++)
                    {
                        Color.Add(Color[0]);
                    }
                }

                bool[] isMasterStory = IsMasterStory.ToArray();
                string[] similarToStory = SimilarToStory.ToArray();
                bool[] spliceAbove = SpliceAbove.ToArray();
                double[] spliceHeight = SpliceHeight.ToArray();
                int[] color = Color.ToArray();

                ret = mySapModel.Story.SetStories_2(BaseElevation, NumberStories, ref storyNames, ref storyHeights, ref isMasterStory, ref similarToStory, ref spliceAbove, ref spliceHeight, ref color);
                
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

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.Edit; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("9934e52c-d75a-4c4a-a80c-ea7fc7f862bd");
    }
}