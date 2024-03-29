using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System;
using System.Collections.Generic;
using ETABSv1;
using System.Windows.Forms;

namespace Tomaso
{
    public class AnalyticalPanelType : GH_Component
    {
        private PanelType panelType = PanelType.Undefined;

        /// <summary>
        /// Panel Type
        /// </summary>
        public AnalyticalPanelType()
          : base("AnalyticalPanelType", "AnalyticalPanelType",
              "Snap Panels",
              "SAM", "Analytical")
        {

        }

        public enum PanelType
        {
            Undefined,
            Wall,
            WallExternal,
            WallInternal,
        }

        /*
         * public override bool Write(GH_IWriter writer)
        {
            writer.SetInt32("PanelType", (int)panelType);
            return base.Write(writer);
        }

        public override bool Read(GH_IReader reader)
        {
            int aIndex = -1;
            if (reader.TryGetInt32("PanelType", ref aIndex))
                panelType = (PanelType)aIndex;

            return base.Read(reader);
        }
        */

        protected override void AppendAdditionalComponentMenuItems(ToolStripDropDown menu)
        {
            foreach (PanelType pt in Enum.GetValues(typeof(PanelType)))
                GH_Component.Menu_AppendItem(menu, pt.ToString(), Menu_PanelTypeChanged, true, pt == this.panelType).Tag = pt;

        }

        private void Menu_PanelTypeChanged(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem item && item.Tag is PanelType)
            {
                //Do something with panelType
                this.panelType = (PanelType)item.Tag;
                ExpireSolution(true);
            }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_InputParamManager inputParamManager)
        {

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_OutputParamManager outputParamManager)
        {
            outputParamManager.AddGenericParameter("PanelType", "PanelType", "Analytical PanelType", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="dataAccess">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess dataAccess)
        {
            dataAccess.SetData(0, panelType);
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("25a6b405-19ab-4ff1-9666-7760997ccfdd"); }
        }
    }
}