using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Collections.Specialized;
using BitCollectors.UIAutomationLib.Entities;
using System;

namespace BitCollectors.UIAutomationLib.Helpers
{
    public class XmlHelper
    {
        private static UIAutomation _uiAutomation = null;

        #region Public Static Methods
        public static UIAutomation ProcessXmlString(string xml)
        {
            return ProcessXmlString(xml, null);
        }

        public static UIAutomation ProcessXmlString(string xml, StringDictionary macros)
        {
            XElement root = XElement.Parse(xml);

            return ProcessXmlElement(root, macros);
        }

        /// <summary>
        /// Process the UI Automation XML configuration information into a usable business object
        /// </summary>
        /// <returns>The business object with the configuration data converted from the XML doc</returns>
        public static UIAutomation ProcessXmlFile(string automationFile)
        {
            return ProcessXmlFile(automationFile, null);
        }

        public static UIAutomation ProcessXmlFile(string automationFile, StringDictionary macros)
        {
            if (File.Exists(automationFile))
            {
                XElement root = XElement.Load(automationFile);

                return ProcessXmlElement(root, macros);
            }

            return _uiAutomation;
        }
        #endregion

        private static UIAutomation ProcessXmlElement(XElement rootNode, StringDictionary macros)
        {
            VerifyMacrosAreValid(macros);

            _uiAutomation = new UIAutomation();
            
            var xmlData = from el in rootNode.Elements()
                          select el;

            foreach (XElement element in xmlData)
            {
                switch (element.Name.LocalName)
                {
                    case "FormHandle":
                        ProcessFormHandleElement(element);
                        break;

                    case "VerifyScreen":
                        ProcessVerifyScreenElement(element);
                        break;

                    case "InputActions":
                        ProcessInputActionsElement(element, macros);
                        break;

                    case "Messages":
                        ProcessMessagesElement(element);
                        break;
                }
            }

            return _uiAutomation;
        }

        #region Private Static Methods
        private static void VerifyMacrosAreValid(StringDictionary macros)
        {
            if (macros != null)
            {
                foreach (string key in macros.Keys)
                {
                    Regex regex = new Regex(@"^\{[A-Za-z0-9]+\}$");
                    if (!regex.IsMatch(key))
                    {
                        throw new Exception("Invalid Macro.  Macro keys should be surrounded with curly braces and only contain alpha-numeric characters.");
                    }
                }
            }
        }

        private static void ProcessFormHandleElement(XElement element)
        {
            if (element.Attribute("TitleBar") != null)
            {
                _uiAutomation.FormHandleTitleBar = element.Attribute("TitleBar").Value;
            }

            if (element.Attribute("ClassName") != null)
            {
                _uiAutomation.FormHandleClassName = element.Attribute("ClassName").Value;
            }

            if (element.Attribute("TitleBarRegex") != null)
            {
                _uiAutomation.FormHandleTitleBarRegex = element.Attribute("TitleBarRegex").Value;
            }
        }

        private static void ProcessVerifyScreenElement(XElement element)
        {
            if (element.Attribute("Type") != null)
            {
                if (element.Attribute("Type").Value == "PixelColor")
                {
                    _uiAutomation.VerifyScreenOptions.VerifyScreenType = VerifyScreenTypes.PixelColor;
                }
                else
                {
                    _uiAutomation.VerifyScreenOptions.VerifyScreenType = VerifyScreenTypes.None;
                }
            }

            if (element.Attribute("X") != null && element.Attribute("Y") != null)
            {
                int x, y;
                if (int.TryParse(element.Attribute("X").Value, out x) && int.TryParse(element.Attribute("Y").Value, out y))
                {
                    Point screenPoint = new Point(x, y);
                    _uiAutomation.VerifyScreenOptions.VerifyScreenPixelLocation = screenPoint;
                }
            }

            if (element.Attribute("PixelColorValue") != null)
            {
                string pixelColorString = element.Attribute("PixelColorValue").Value.Trim();

                if (pixelColorString.StartsWith("#"))
                {
                    _uiAutomation.VerifyScreenOptions.VerifyScreenPixelColor =
                        ColorTranslator.FromHtml(pixelColorString);
                }
                else
                {
                    Regex regex = new Regex(@"^\{([0-9]{1,3})\,([0-9]{1,3})\,([0-9]{1,3})\}$");
                    Match match = regex.Match(pixelColorString);
                    if (match.Groups.Count > 3)
                    {
                        int colorR = int.Parse(match.Groups[1].Value);
                        int colorG = int.Parse(match.Groups[2].Value);
                        int colorB = int.Parse(match.Groups[3].Value);

                        _uiAutomation.VerifyScreenOptions.VerifyScreenPixelColor = Color.FromArgb(
                            colorR,
                            colorG,
                            colorB);
                    }
                }
            }
        }

        private static void ProcessInputActionsElement(XElement element, StringDictionary macros)
        {
            var inputBatchData = from batchEl in element.Elements()
                                 where batchEl.Name.LocalName == "InputBatch"
                                 select batchEl;

            foreach (XElement batchElement in inputBatchData)
            {
                InputBatch batch = new InputBatch();

                if (batchElement.Attribute("Timeout") != null)
                    batch.Timeout = int.Parse(batchElement.Attribute("Timeout").Value);

                var keyStrokeData = from strokeEl in batchElement.Elements()
                                    select strokeEl;

                foreach (XElement strokeElement in keyStrokeData)
                {
                    switch (strokeElement.Name.LocalName)
                    {
                        case "KeyStroke":
                            string strokeType = "downup";
                            string strokeValue = null;
                            //List<string> keys = new List<string>();
                            string[] keys = new string[1];
                            int strokeRepeat = 1;

                            if (strokeElement.Attribute("Type") != null)
                            {
                                strokeType = strokeElement.Attribute("Type").Value.ToLower();
                            }

                            if (strokeElement.Attribute("Value") != null)
                            {
                                strokeValue = strokeElement.Attribute("Value").Value.ToLower();

                                if (strokeValue.Contains("+"))
                                {
                                    keys = strokeValue.Split('+');
                                }
                                else if (strokeValue.StartsWith("{") && strokeValue.EndsWith("}"))
                                {
                                    keys[0] = strokeValue;
                                }
                                else
                                {
                                    keys = strokeValue.ToCharArray().Select(c => c.ToString()).ToArray();
                                }
                            }

                            if (strokeElement.Attribute("Repeat") != null)
                            {
                                strokeRepeat = int.Parse(strokeElement.Attribute("Repeat").Value);
                            }

                            if (macros != null && macros.ContainsKey(strokeValue))
                            {
                                char[] keyStrokeChars = macros[strokeValue].ToCharArray();

                                foreach (char keyStrokeChar in keyStrokeChars)
                                {
                                    KeyStroke macroKeystrokeDown = new KeyStroke();
                                    macroKeystrokeDown.Type = KeyStrokeTypes.Down;
                                    macroKeystrokeDown.Value = ConvertKeyStrokeToCode(keyStrokeChar.ToString());
                                    batch.InputActions.Add(macroKeystrokeDown);

                                    KeyStroke macroKeystrokeUp = new KeyStroke();
                                    macroKeystrokeUp.Type = KeyStrokeTypes.Up;
                                    macroKeystrokeUp.Value = ConvertKeyStrokeToCode(keyStrokeChar.ToString());
                                    batch.InputActions.Add(macroKeystrokeUp);
                                }
                            }
                            else
                            {
                                for (int i = 1; i <= strokeRepeat; i++)
                                {
                                    if (strokeType == "down" || strokeType == "downup")
                                    {
                                        for (int j = 0; j < keys.Length; j++)
                                        {
                                            KeyStroke keyStrokeDown = new KeyStroke();
                                            keyStrokeDown.Type = KeyStrokeTypes.Down;
                                            keyStrokeDown.Value = ConvertKeyStrokeToCode(keys[j]);
                                            batch.InputActions.Add(keyStrokeDown);
                                        }
                                    }

                                    if (strokeType == "up" || strokeType == "downup")
                                    {
                                        for (int j = keys.Length - 1; j >= 0; j--)
                                        {
                                            KeyStroke keyStrokeUp = new KeyStroke();
                                            keyStrokeUp.Type = KeyStrokeTypes.Up;
                                            keyStrokeUp.Value = ConvertKeyStrokeToCode(keys[j]);
                                            batch.InputActions.Add(keyStrokeUp);
                                        }
                                    }
                                }
                            }

                            break;

                        case "MouseClick":
                            string clickType = "click";
                            int clickRepeat = 1;
                            int x = 0;
                            int y = 0;
                            string relativeTo = "window";

                            if (strokeElement.Attribute("Type") != null)
                            {
                                clickType = strokeElement.Attribute("Type").Value.ToLower();
                            }

                            if (strokeElement.Attribute("RelativeTo") != null)
                            {
                                relativeTo = strokeElement.Attribute("RelativeTo").Value.ToLower();
                            }

                            if (strokeElement.Attribute("Repeat") != null)
                            {
                                clickRepeat = int.Parse(strokeElement.Attribute("Repeat").Value);
                            }

                            if (strokeElement.Attribute("X") != null)
                            {
                                x = int.Parse(strokeElement.Attribute("X").Value);
                            }

                            if (strokeElement.Attribute("Y") != null)
                            {
                                y = int.Parse(strokeElement.Attribute("Y").Value);
                            }

                            for (int i = 1; i <= clickRepeat; i++)
                            {
                                MouseClick mouseClickMoveTo = new MouseClick();
                                mouseClickMoveTo.Type = MouseClickTypes.Move;
                                mouseClickMoveTo.Coordinates = new Point(x, y);
                                mouseClickMoveTo.RelativeTo = relativeTo;
                                batch.InputActions.Add(mouseClickMoveTo);

                                if (clickType == "down" || clickType == "click")
                                {
                                    MouseClick mouseClick = new MouseClick();
                                    mouseClick.Type = MouseClickTypes.Down;
                                    mouseClick.Coordinates = new Point(x, y);
                                    mouseClick.RelativeTo = relativeTo;
                                    batch.InputActions.Add(mouseClick);
                                }

                                if (clickType == "up" || clickType == "click")
                                {
                                    MouseClick mouseClick = new MouseClick();
                                    mouseClick.Type = MouseClickTypes.Up;
                                    mouseClick.Coordinates = new Point(x, y);
                                    mouseClick.RelativeTo = relativeTo;
                                    batch.InputActions.Add(mouseClick);
                                }
                            }
                            break;
                    }


                }

                _uiAutomation.InputBatches.Add(batch);
            }
        }

        private static void ProcessMessagesElement(XElement element)
        {
            var messageData = from messageEl in element.Elements()
                              select messageEl;

            foreach (XElement messageElement in messageData)
            {
                if (messageElement.Attribute("Type") != null && messageElement.Attribute("Value") != null)
                {
                    switch (messageElement.Attribute("Type").Value)
                    {
                        case "WindowHandleFailed":
                            _uiAutomation.WindowHandleFailedMessage =
                                messageElement.Attribute("Value").Value;
                            break;

                        case "VerifyScreenFailed":
                            _uiAutomation.VerifyScreenFailedMessage =
                                messageElement.Attribute("Value").Value;
                            break;
                    }
                }
            }
        }
        #endregion

        private static short ConvertKeyStrokeToCode(string magicString)
        {
            if (magicString.Length == 1)
            {
                return Win32Helper.GetCharacterValue(magicString.ToLower()[0]);
            }
            else
            {
                switch (magicString.ToUpper().Trim())
                {
                    case "{CTRL}":
                    case "{CONTROL}":
                        return (short)Win32Helper.VirtualKeys.CONTROL;

                    case "{SHIFT}":
                        return (short)Win32Helper.VirtualKeys.SHIFT;

                    case "{ALT}":
                        return (short)Win32Helper.VirtualKeys.MENU;

                    case "{ENTER}":
                    case "{RETURN}":
                        return (short)Win32Helper.VirtualKeys.RETURN;

                    case "{TAB}":
                        return (short)Win32Helper.VirtualKeys.TAB;

                    case "{HOME}":
                        return (short)Win32Helper.VirtualKeys.HOME;

                    case "{END}":
                        return (short)Win32Helper.VirtualKeys.END;

                    case "{DEL}":
                    case "{DELETE}":
                        return (short)Win32Helper.VirtualKeys.DELETE;

                    case "{LWIN}":
                        return (short)Win32Helper.VirtualKeys.LWIN;

                    case "{RWIN}":
                        return (short)Win32Helper.VirtualKeys.RWIN;

                    case "{LEFT}":
                        return (short)Win32Helper.VirtualKeys.LEFT;

                    case "{RIGHT}":
                        return (short)Win32Helper.VirtualKeys.RIGHT;

                    case "{UP}":
                        return (short)Win32Helper.VirtualKeys.UP;

                    case "{DOWN}":
                        return (short)Win32Helper.VirtualKeys.DOWN;

                    case "{ESC}":
                    case "{ESCAPE}":
                        return (short)Win32Helper.VirtualKeys.ESCAPE;

                    default:
                        return Win32Helper.GetCharacterValue('Z');
                }
            }
        }
    }
}
