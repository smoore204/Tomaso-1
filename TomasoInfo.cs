using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace Tomaso
{
    public class TomasoInfo : GH_AssemblyInfo
    {
        public override string Name => "Tomaso";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "";

        public override Guid Id => new Guid("1D6186C6-CDA2-4376-830E-93FD0AA9EF2D");

        //Return a string identifying you or your company.
        public override string AuthorName => "Sophie Moore";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "soph.x.moore@gmail.com";
    }
}