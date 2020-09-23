<div align="center">

## LaVolpe Desktop Icon Tweaks \(Finished\)

<img src="PIC2006520182425988.gif">
</div>

### Description

Desktop subclassing with VB. An icon uber-tweaker. FINAL CUT barring bugs. Per-Icon color &amp; caption settings transforms a blah desktop to something pretty nice. Other options and tools are included with the project.

----

This project put together to highlight the flexibility of VB. The effects are accomplished by subclassing the desktop. The DLL used to subclass the desktop is created in VB, the DLL used for injection into the desktop (global hooking) is the same DLL.

----

Well, I can't upload compiled DLLs on PSC, you will have to compile the hybrid DLL yourself. I have included instructions on 2 ways to accomplish this: an easy way &amp; the hard way (my way). Unzip the file using "Use Folders" option because this is 3 projects in one.

----

Updates: 15May05::Fix issues reported by Steve and others; DLL was not coloring text &amp; was result of DLL typo. 17May05::added save/restore/un-restore icon positions, tweaked GUI, &amp; enabled removal of injected DLL. 17May05::found/fixed bug that can crash when items sent to recycle bin. InstallDLL function now does version check before install. Fixed losing saved icon positions on startup &amp; ability to toggle AutoArrange on NT &amp; 9x. 20May05::Reworked RebuildListView to prevent NT icons from unrestoring in some cases, added more options, tweaked GUI. 24May05:: Fixed auto-icon restore after display res change &amp; a couple minor bugs.

----

Always destroy your previous version of this project &amp; run the TxtToBin to create compiled DLLs or compile them yourself. As of this upload, the frmDTop will have the required DLL version nr to help reduce any confusion. This version &amp; any future updates will have 2 DLLs. One is simply used to remove the injected DLL. Read comments in Declaration section of frmDTop &amp; also the HowToSetup.rtf file.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2006-05-22 09:46:44
**By**             |[LaVolpe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lavolpe.md)
**Level**          |Advanced
**User Rating**    |5.0 (125 globes from 25 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[LaVolpe\_De1996435232006\.zip](https://github.com/Planet-Source-Code/lavolpe-lavolpe-desktop-icon-tweaks-finished__1-65301/archive/master.zip)

### API Declarations

Dozens





