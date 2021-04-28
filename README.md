# IguanaTex

(C) Jonathan Le Roux and Zvika Ben-Haim (Windows), Tsung-Ju Chiang and Jonathan Le Roux (Mac)

Website: http://www.jonathanleroux.org/software/iguanatex/

IguanaTex is a PowerPoint add-in which allows you to insert LaTeX equations into your PowerPoint presentation on Windows and Mac. It is distributed completely for free, along with its source code.

This repository hosts the source code in a form that can be easily tracked, shared, and discussed (the VBA code is exported using the [ExportVBA macro](https://github.com/Jonathan-LeRoux/IguanaTex/blob/master/ExportVBA.bas)).


## System Requirements
### Windows

* OS: Windows 2000 or later (32- or 64-bit).
* PowerPoint: 
  * IguanaTex has been tested with PowerPoint 2010, 2013, 2016, 2019 (both 32 and 64 bit), Office 365, and PowerPoint 2003. It is likely to also work in PowerPoint 2000 and 2007.
  * SVG support is only available for Office 365 (and PowerPoint 2019? To be confirmed)
* LaTeX: [TeXLive](https://www.tug.org/texlive/) or [MiKTeX](http://miktex.org/)
* [GhostScript](http://www.ghostscript.com/download/gsdnld.html) (if the latest version raises issues, try gs9.26)
* [ImageMagick](http://www.imagemagick.org/script/download.php#windows) (check the box to "Install development headers and libraries for C and C++" or, in recent versions, to "Install legacy utilities (e.g., convert)")
* (Optional) [TeX2img](https://github.com/abenori/TeX2img), used for vector graphics output via EMF ([Download](https://www.ms.u-tokyo.ac.jp/~abenori/soft/index.html#TEX2IMG)).

### Mac

* Intel or Apple Silicon Mac
  * On Apple Silicon based Macs, PowerPoint needs to be run with the setting "Open using Rosetta."
* PowerPoint for Mac: 
  * Powerpoint 2019 or Office 365
  * PowerPoint 2016 (Version 16.16.7 190210 or later) (to be confirmed; some features, e.g., SVG, may not work)
* [MacTeX](https://www.tug.org/mactex/)


## Download and Install

### Windows 

1. **Download the .ppam add-in** file from this repository's [Releases page](https://github.com/Jonathan-LeRoux/IguanaTex), or from the [IguanaTex Download page](http://www.jonathanleroux.org/software/iguanatex/download.html).
2. **Load the add-in**: in "File" > "Options" > "Add-Ins" > "Manage:", choose "PowerPoint Add-Ins" then "Go...", then click  "Add New", select the .ppam file in the folder where you downloaded it, then "Close" (if you downloaded the .pptm source and saved it as .ppam, it will be in the default Add-In folder).
3. **Create and set a temporary file folder**: IguanaTex needs access to a folder with read/write permissions to store temporary files.
  * The default is "C:\Temp\". If you have write permissions under "C:\", create the folder "C:\Temp\". You're all set.
  * If you cannot create this folder, choose or create a folder with write permission at any other location. In the IguanaTex tab, choose "Main Settings" and put the path to the folder of your choice. You can also use a relative path under the presentation's folder (e.g., ".\" for the presentation folder itself).
4. **Install and set path to GhostScript and ImageMagick**:
  * Set the **full** path to gswin32c.exe or gswin64c.exe (note the "c"!) and to ImageMagick's convert.exe in the "Main Settings" window.
  * Best way to make sure the path is correct is to use the "..." button next to each path and navigate to the correct file.
5. (Optional) **Install and set path to TeX2img**:
  * Only needed for vector graphics support via EMF (compared to SVG, pros: available on all PowerPoint versions, fully modifiable shapes; cons: some displays randomly suffer from distortions)
  * Download from [this link](https://www.ms.u-tokyo.ac.jp/~abenori/soft/index.html#TEX2IMG) (more details on TeX2img on their [Github repo](https://github.com/abenori/TeX2img))
  * After unpacking TeX2img somewhere on your machine, run TeX2img.exe once to let it automatically set the various paths to latex/ghostscript, then set the **full** path to TeX2imgc.exe (note the "c"!) in the "Main Settings" window.
6. (Optional) **Install LaTeXiT-metadata**:
  * Needed to convert displays generated with LaTeXiT on Mac into IguanaTex displays
  * Stay tuned for the release (code is ready, just need to figure out how best to release it)

**Other settings**:
* If you have a non-standard LaTeX installation, you can specify the folder in which the executables are included.
* If you would like to have the option of using an external editor, e.g., when debugging LaTeX source code, you can specify the path to that editor. If you would like to use that editor by default over the IguanaTex edit window, check the "use as default" checkbox.

### Mac

1. **Download the "prebuilt files for Mac" zip** from this repository's [Releases page](https://github.com/Jonathan-LeRoux/IguanaTex/Releases)  
There are 3 files to install:
* `IguanaTex.scpt`: AppleScript file for handling file and folder access
* `libIguanaTexHelper.dylib`: library for creating native text views; source code included in the git repo, under "IguanaTexHelper/"
* `IguanaTexMac.ppam`: main add-in file

2. **Install `IguanaTex.scpt`**
```bash
mkdir -p ~/Library/Application\ Scripts/com.microsoft.Powerpoint
cp ./IguanaTex.scpt ~/Library/Application\ Scripts/com.microsoft.Powerpoint/IguanaTex.scpt
```

3. **Install `libIguanaTexHelper.dylib`**
```bash
sudo mkdir -p '/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized'
sudo cp ./libIguanaTexHelper.dylib '/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/libIguanaTexHelper.dylib'
```

4. **Load the add-in**: Start PowerPoint. From the menu bar, select Tools > PowerPoint Add-ins... > '+' , and choose `IguanaTexMac.ppam`

5. **Verify that paths are set correctly**: 
  * Click on "Main Settings" in the IguanaTex ribbon tab, and verify that the paths to GhostScript, ImageMagick, LaTeX, and libgs.9.dylib are set correctly by clicking on each "..." button next to them.
  * If you cannot find them, open a terminal and use "locate \<filename\>".

6. (Optional) **Install LaTeXiT-metadata**:
  * Needed to convert displays generated with LaTeXiT into IguanaTex displays
  * Stay tuned for the release (code is ready, just need to figure out how best to release it)
