using System.Collections.Generic;

namespace BitCollectors.UIAutomationLib.Entities
{
    public class UIAutomation
    {
        #region Private Fields
        private string _windowHandleFailedMessage =
            "Could not find the application running.  Please start the application or restart it and try again.";

        private string _verifyScreenFailedMessage =
            "";
        #endregion

        #region Public Properties
        public string FormHandleTitleBar { get; set; }

        public string FormHandleClassName { get; set; }

        public string FormHandleTitleBarRegex { get; set; }

        public VerifyScreen VerifyScreenOptions { get; set; }

        public List<InputBatch> InputBatches { get; set; }

        public string WindowHandleFailedMessage
        {
            get { return _windowHandleFailedMessage; }
            set { _windowHandleFailedMessage = value; }
        }

        public string VerifyScreenFailedMessage
        {
            get { return _verifyScreenFailedMessage; }
            set { _verifyScreenFailedMessage = value; }
        }
        #endregion

        public UIAutomation()
        {
            this.VerifyScreenOptions = new VerifyScreen();
            this.InputBatches = new List<InputBatch>();
        }
    }
}
