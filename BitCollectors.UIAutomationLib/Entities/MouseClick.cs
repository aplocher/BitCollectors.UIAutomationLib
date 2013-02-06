using BitCollectors.UIAutomationLib.Interfaces;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace BitCollectors.UIAutomationLib.Entities
{
    public class MouseClick: IInputAction
    {
        public MouseClickTypes Type { get; set; }

        public Point Coordinates { get; set; }

        public string RelativeTo { get; set; }

        public InputActionTypes InputAction
        {
            get
            {
                return InputActionTypes.MouseClick;
            }
        }
    }
}
