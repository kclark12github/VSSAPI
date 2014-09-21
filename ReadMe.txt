**********************************************************
IMPORTANT INFORMATION ABOUT THE VSS OLE Automation Sample 
**********************************************************

This sample demonstrates the Visual SourceSafe 6.0 OLE Automation interface.
It includes the DiffMerge ActiveX control which implements a version of
the Visual SourceSafe View, Difference and Visual Merge Dialogs. Please note 
that there is no support for this sample or the DiffMerge control. 

-----------------------------------------------------------------------------
Setup:
-----------------------------------------------------------------------------

To run this sample you must have a licensed copy of Visual SourceSafe
6.0 and Visual Basic 5.0 (SP3) installed on your system. This sample program 
requires certain files that ship only with Visual SourceSafe 6.0.

You may run this sample from either the Visual Basic development environment
or by running the file VSSOLESample.exe. To run the executable you must have 
the Visual Basic 5.0 (SP3) run time libraries installed.

The following steps are required for this sample to run properly:

1. Move the file DiffMergeCtl.ocx  (included in the Zip file) to the Win32 
   folder of your local Visual SourceSafe 6.0 installation. 

2. Register the file <VSS Path>\Win32\DiffMergeCtl.ocx by running:

   Regsvr32.exe <VSS Path>\Win32\DiffMergeCtl.ocx

   where <VSS Path> is the path to your local installation of SourceSafe 6.0.

-----------------------------------------------------------------------------
Notes:
-----------------------------------------------------------------------------

This sample reads and uses a variety of "personal settings" from the user's 
SS.INI file. These settings include:

   Warning Dialog Settings. This includes warnings for:

	Delete (file or project)
	Exit SourceSafe
	Destroy (file or project)
	Purge (file or project)
	Checkout (an already checked out file)
	UndoCheckOut (a file that has changed)

  Double Click on a file (edit or view)
  Difference Format (for displaying differences-DiffMergeControl only)

When the sample program launches it reads and uses these settings from the 
user's SS.INI file. You may modify these settings by selecting the Options 
command from the Tools menu of the sample. Changing these settings will not 
write back to the user's SS.INI file.

You may choose to disable use of the DiffMergeControl by choosing the Options 
command from the Tools menu and disabling the "Use OCX control" CheckBox from 
the "View\Editing files" and "Showing File Difference" sections of the 
View\Diff Tab. Disabling these will also disable use of the Visual Merge dialog.

Check the Visual SourceSafe web site for additional documentation on the OLE 
Automation interface.
 
