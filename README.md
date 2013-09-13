FreeMiCal
=========

FreeMiCal allows the bulk export of MS Outlook appointment items to RFC 2445 conformant iCal format. .NET 2.0 or greater and Outlook 2003 or 2007 Interop Services required. Download zip for deployable executable.


FreeMiCal Windows version
=========================

FreeMiCal.exe requires fMiCal.exe in the same directory to export
Outlook 2007 calendar items.

It will set the export parameters for fMiCal.exe (the command line
version) and then call fMiCal as a process with a hidden window.

How to run:
-----------

Double click FreeMiCal.exe to open the graphical user interface.


Command line parameters of fMiCal.exe
-------------------------------------

usage: fmical [option|...]

 --help print this page. Same as ?
 --start=n start export with record n
 --end=m end export after record m. Negative numbers are ignored
 --output=file Filename to export calendar items
 --profile=name Outlook profile name

 parameters can be abbreviated windows (e.g. /h) or old unix style (e.g. -h)
 equal signs (=) may be replaced with colons (:)
  

Implementation details
----------------------

FreeMiCal.exe queries Microsoft Outlook for the number of calendar items.
If you choose to export the calendar (by clicking "Free them..."),
FreeMiCal spawns fMiCal as a hidden process.

fMiCal reports its export status in a synchronization file fMiCal.run.

FreeMiCal monitors this file and updates the progress bar accordingly.

You can choose to stop the export by clicking "Cancel...". This stops
fMiCal.exe.

The synchronization file is created by fMiCal.exe and deleted by FreeMiCal.exe.


Acknowlegement
--------------

Thanks to Robert John / Microsoft Austria for giving the final indication about
how to be independent of versions of Outlook.


@2007 by RSB
