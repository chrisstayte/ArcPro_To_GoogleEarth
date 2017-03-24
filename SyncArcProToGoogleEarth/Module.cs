using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using System.Threading.Tasks;

namespace SyncArcProToGoogleEarth
{
    internal class Module : ArcGIS.Desktop.Framework.Contracts.Module
    {
        private static Module _this = null;

        /// <summary>
        /// Retrieve the singleton instance to this module here
        /// </summary>
        public static Module Current
        {
            get
            {
                return _this ?? (_this = (Module)FrameworkApplication.FindModule("SyncArcProToGoogleEarth_Module"));
            }
        }

        #region Overrides
        /// <summary>
        /// Called by Framework when ArcGIS Pro is closing
        /// </summary>
        /// <returns>False to prevent Pro from closing, otherwise True</returns>
        protected override bool CanUnload()
        {
            //TODO - add your business logic
            //return false to ~cancel~ Application close
            return true;
        }

        #endregion Overrides

    }
}
