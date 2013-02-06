using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BitCollectors.UIAutomationLib
{
    public enum InputActionTypes
    {
        KeyStroke,
        MouseClick
    }

    public enum KeyStrokeTypes
    {
        Down,
        Up,
        DownUp
    }

    public enum MouseClickTypes
    {
        Move,
        Click,
        Down,
        Up
    }
}
