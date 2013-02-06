using BitCollectors.UIAutomationLib.Interfaces;

namespace BitCollectors.UIAutomationLib.Entities
{
    public class KeyStroke: IInputAction
    {
        public KeyStrokeTypes Type { get; set; }

        public int Repeat { get; set; }

        public short Value { get; set; }

        public InputActionTypes InputAction
        {
            get
            {
                return InputActionTypes.KeyStroke;
            }
        }
    }
}
