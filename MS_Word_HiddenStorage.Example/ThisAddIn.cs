﻿using System;

namespace OfficeTools.Examples {
    public partial class ThisAddIn {
        void ThisAddIn_Startup(object sender, EventArgs e) { }

        void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        void InternalStartup() {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}