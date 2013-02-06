using System.Drawing;

namespace BitCollectors.UIAutomationLib.Entities
{
    public enum VerifyScreenTypes
    {
        None,
        PixelColor
    }

    public class VerifyScreen
    {
        public VerifyScreenTypes VerifyScreenType { get; set; }

        public Point VerifyScreenPixelLocation { get; set; }

        public Color VerifyScreenPixelColor { get; set; }
    }
}
