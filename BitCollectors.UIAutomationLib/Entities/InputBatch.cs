using BitCollectors.UIAutomationLib.Interfaces;
using System.Collections.Generic;

namespace BitCollectors.UIAutomationLib.Entities
{
    public class InputBatch
    {
        /// <summary>
        /// Set of Keystrokes to execute in this batch.
        /// </summary>
        public List<IInputAction> InputActions { get; set; }

        /// <summary>
        /// Timeout between keystroke batches in milliseconds.
        /// </summary>
        public int Timeout { get; set; }

        public InputBatch()
        {
            this.InputActions = new List<IInputAction>();
        }
    }
}
