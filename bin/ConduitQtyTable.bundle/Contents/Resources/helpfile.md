#Conduit Quantity Table helpfile

###Contents
<ul>
	<li><a href="#description">Description</a></li>
	<li><a href="#install">Installation</a></li>
	<li><a href="#commands">Commands Available</a></li>
	<li><a href="#troubleshoot">Troubleshooting</a></li>
	<li><a href="#uninstall">Uninstall</a></li>
</ul>

<hr>

<a name="desctription"></a>
####Description

This plugin works for AutoCAD MEP 2012 and above. It allows a user to select multiple conduits and create an Mleader object with a Mtext attachment that is populated with the Quantity and size of each conduit in the selection set.

User can also set the default layer for the Mleader to be created on (Default set to Layer '0').

<a name="install"></a>
####Installation

To install the plugin, run the <pre>ConduitQtyTable.exe</pre> file. This will install into your ApplicationPlugins directory.

<a name="commands"></a>
####Commands Available

<pre>ConduitQtyTable</pre> - This is the main command to initiate the tool. Follow the prompts on the command line to execute steps.

<pre>SetMleaderLayer</pre> - This sets the default layer to place the final objects on.

<a name="troubleshoot"></a>
####Troubleshooting

If you have any problems with the program, please report them to <a href="mailto:kmorin@dyna-sd.com">kmorin@dyna-sd.com</a>

<a name="uninstall"></a>
####Uninstall

If you wish to uninstall the application, simply remove the folder at this location:

<em>C:\Program Files\Autodesk\ApplicationPlugins\ConduitQtyTable.bundle</em>