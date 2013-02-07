BitCollectors.UIAutomationLib
=============================

This automation library allows for the simulating of keystrokes and mouse clicks.  There are other solutions which may be better
suited for native Windows apps, such as the .NET UI Automation framework, which will inject windows messages directly in order 
to populate text areas, instead of simulating key strokes.

This library works with just about every system I've needed to automate, including systems that run through Terminal Services or
Citrix.

I've included the ability to define and use macros.  A macro consists of a StringDictionary that gives you the ability to push
a batch of key strokes out from some user input.  For example, in your Automation Script (XML file), you could define a macro
called {USERTEXT} and then associate that to some text from a textbox.  When your script runs, it will generate key strokes for
that text.  You can have as many macros as you want defined.

NOTE: I've had issues with certain Windows API calls on X64 machines when the build platform is configured for Any CPU.  If you
run in to problems with the automation not working, try setting your build platform to X86.

To get started, reference BitCollectors.UIAutomationLib.dll in your .NET Project.  The code to harness the DLL is very easy to
implement.  See below:

Basic usage (C#):

    using BitCollectors.UIAutomationLib.Entities;
    using BitCollectors.UIAutomationLib.Helpers;

    ...

    UIAutomation automationConfig = XmlHelper.ProcessXmlFile("C:\\Temp\\MyAutomationScript.xml");

    if (automationConfig != null)
    {
        Win32Helper win32 = new Win32Helper();
        win32.ExecuteAutomation(automationConfig);
    }

Basic usage with macros (C#):

    using BitCollectors.UIAutomationLib.Entities;
    using BitCollectors.UIAutomationLib.Helpers;
    using System.Collections.Specialized;

    ...

    StringDictionary macros = new StringDictionary();
    macros.Add("{USERTEXT}", textBox1.Text);

    UIAutomation automationConfig = XmlHelper.ProcessXmlFile("C:\\Temp\\MyAutomationScript.xml", macros);

    if (automationConfig != null)
    {
        Win32Helper win32 = new Win32Helper();
        win32.ExecuteAutomation(automationConfig);
    }

In addition to XmlHelper.ProcessXmlFile, the XmlHelper class contains a static method called ProcessXmlString.  This is used to 
pass in a raw XML string, and not rely on a file on the file system.  This could be helpful if you want to programatically 
generate an XML Script and send it in to simulate keystrokes/mouse clicks.


Built in keystroke keywords include:

	{CTRL} or {CONTROL}
	{SHIFT}
	{ALT}
	{ENTER} or {RETURN}
	{TAB}
	{HOME}
	{END}
	{DEL} or {DELETE}
	{LWIN}
	{RWIN}
	{ESC} or {ESCAPE}
	{UP}
	{DOWN}
	{LEFT}
	{RIGHT}


The following XML example is located at BitCollectors.UIAutomationLib\Config\TestAutomation.xml in the solution. 

	<?xml version="1.0" encoding="utf-8" ?>
	<Automation>
	  <!-- Find the first open Notepad window -->
	  <FormHandle TitleBarRegex="^(.*?) - Notepad$" />

	  <InputActions>
		<InputBatch Timeout="150">
		  <!-- Click in the window -->
		  <MouseClick Type="Click" RelativeTo="Window" X="80" Y="80" />
		</InputBatch>

		<InputBatch Timeout="200">
		  <!-- Press CTRL+A to select all text, and then press DELETE to clear out the window -->
		  <KeyStroke Type="downup" Value="{CTRL}+A" />
		  <KeyStroke Type="downup" Value="{DEL}" />
		</InputBatch>

		<InputBatch Timeout="10">
		  <!-- Send in the {USERTEXT} macro.  This will simulate keystrokes for whatever you define {USERTEXT} to be -->
		  <KeyStroke Type="downup" Value="{USERTEXT}" />
		</InputBatch>

		<InputBatch Timeout="10">
		  <!-- Press the ENTER key to start a new line -->
		  <KeyStroke Type="downup" Value="{ENTER}" />
		</InputBatch>

		<InputBatch Timeout="10">
		  <!-- This will simulate uppercase letters vs lowercase letters by holding the SHIFT key down -->
		  <!-- Hold down the SHIFT key and type C A P S, release the SHIFT key, type a SPACE followed by the letters L O W E R -->
		  <KeyStroke Type="down" Value="{SHIFT}" />
		  <KeyStroke Type="downup" Value="C" />
		  <KeyStroke Type="downup" Value="A" />
		  <KeyStroke Type="downup" Value="P" />
		  <KeyStroke Type="downup" Value="S" />
		  <KeyStroke Type="up" Value="{SHIFT}" />

		  <KeyStroke Type="downup" Value=" " />
		  <KeyStroke Type="downup" Value="l" />
		  <KeyStroke Type="downup" Value="o" />
		  <KeyStroke Type="downup" Value="w" />
		  <KeyStroke Type="downup" Value="e" />
		  <KeyStroke Type="downup" Value="r" />
		</InputBatch>
	  </InputActions>
	  <Messages>
		<Message Type="WindowHandleFailed" Value="Could not find the application running.  Please make sure application is running." />
	  </Messages>
	</Automation>


XML Tag Information:

	<FormHandle> 
		FormHandle lets you define which application form you want to focus on before sending keystrokes and mouse clicks.
		
		Attributes (you are required to choose one):
			ClassName		- This is the Class Name of the form you're interfacing with
			TitleBar		- The text in the title bar of the form you're interfacing with (for example: "Untitled.txt - Notepad")
			TitleBarRegex	- A regex string that let's you do advanced matches on the Title Bar text (for example: "^(.*?) - Notepad$")

	<InputActions>
		A list of InputBatch elements

	<InputBatch>
		A set of key strokes or mouse clicks to execute.  Child elements are <KeyStroke> or <MouseClick>

		Attributes:
			Timeout	- The number of milliseconds between each batch of commands

	<KeyStroke>
		Attributes:
			Type    - Options: downup (default), down, up. 
			          Typically you would set this to "downup" send a full key stroke (Key Down, Key Up), but if you need to hold a key down, you 
					  would set it to "down".  If you use Type="down", don't forget to set Type="up" before your execution is complete.  The 
					  sample XML above has an example of when you might want to use Type="down" and Type="up".  In the example, we hold down the
					  SHIFT key and send several keystrokes in order to simulate uppercase letters.
			Value	- Options: a single character, a keystroke keyword, a macro definition, or any combination with a '+' sign
					  A single character: "a" or "1" would simulate that key stroke.
					  A keystroke keyword: "{ALT}" or "{TAB}" (see a list of all keystroke keywords above)
					  A macro definition: A user defined input string that is surrounded by curly braces.  See above for more information.
					  A combination with '+' sign: "{CTRL}+a" to simulate holding down CTRL and press the 'A' key
			Repeat	- Number of times to repeat the keystroke. Default 1

	<MouseClick>
		Attributes:
			RelativeTo	- Options: window (default), screen
			X			- X coordinates to put the mouse cursor at.  Relative to the window or screen
			Y			- Y Coordinates to put the mouse cursor at.  Relative to the window or screen
			Type		- Options: click (default), down, up
			Repeat		- Number of times to repeat the click.  Default 1
