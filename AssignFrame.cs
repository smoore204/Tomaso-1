using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;
using System.Windows.Forms;


namespace Tomaso
{
    public class AssignFrame : GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public AssignFrame()
          : base("Assign > Frame", "AssgnFrm",
            "ETABSv1.SapModel.FrameObj\n",
            "Tomaso", "06_Assign")
        {
        }

        public override GH_Exposure Exposure
        {
            get { return GH_Exposure.secondary; }
        }

        protected override void AppendAdditionalComponentMenuItems(ToolStripDropDown menu)
        {
            foreach (string items in MenuItems())
            {
                var tsmi = Menu_AppendItem(menu, items, MenuItemChanged);
                tsmi.ToolTipText = items.ToString();
            }
        }

        private List<string> MenuItems()
        {
            List<string> menuItems = new List<string>();
            menuItems.Add("Additional Mass");
            menuItems.Add("Auto Mesh Options");
            menuItems.Add("Column Splice Overwrite");

            return menuItems;
        }

        private void MenuItemChanged(object sender, EventArgs e)
        {
            AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, sender.ToString());
            if (sender.ToString() == "Additional Mass")
            {
                Params.RegisterInputParam(Naame);
                Params.UnregisterInputParameter(Run);
                ExpireSolution(true);
            }
            if (sender.ToString() == "Auto Mesh Options")
            {
                Params.RegisterInputParam(Run);
                Params.UnregisterInputParameter(Naame);
                ExpireSolution(true);
            }
            if (sender.ToString() == "Column Splice Overwrite")
            {
                Params.RegisterInputParam(Run);
                Params.RegisterInputParam(Run);
                ExpireSolution(true);
            }

        }
        
        Grasshopper.Kernel.Parameters.Param_Boolean Run = new Grasshopper.Kernel.Parameters.Param_Boolean();
        Grasshopper.Kernel.Parameters.Param_String Naame = new Grasshopper.Kernel.Parameters.Param_String();

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            
            {
                Run.NickName = "R";
                Run.Name = "Run";
                Run.Description = "Set to true to run";
                Run.Access = GH_ParamAccess.item;
            }
            {
                Naame.NickName = "N";
                Naame.Name = "Name";
                Naame.Description = "Input Name as string dummy";
                Naame.Access = GH_ParamAccess.item;
            }
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

        }

        /// <summary>
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// You can add image files to your project resources and access them like this:
        /// return Resources.IconForThisComponent;
        /// </summary>

        protected override System.Drawing.Bitmap Icon { get { return Tomaso.Properties.Resources.AFrameSection; } }

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("4bdca759-8fcc-4baf-9943-3eb73a32e58e");
    }
}
