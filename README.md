# IguanaTex

(C) [Jonathan Le Roux](https://www.jonathanleroux.org/) and Zvika Ben-Haim (Windows), [Tsung-Ju Chiang](https://github.com/tsung-ju) and Jonathan Le Roux (Mac)

Website: https://www.jonathanleroux.org/software/iguanatex/

IguanaTex is a PowerPoint add-in which allows you to insert LaTeX equations into your PowerPoint presentation on Windows and Mac. It is distributed completely for free, along with its source code.

This repository hosts the source code in a form that can be easily tracked, shared, and discussed (the VBA code is exported using the [ExportVBA macro](https://github.com/Jonathan-LeRoux/IguanaTex/blob/master/ExportVBA.bas)).


## System Requirements
### Windows

* OS: Windows 2000 or later (32- or 64-bit).
* PowerPoint: 
  * IguanaTex has been tested with PowerPoint 2010, 2013, 2016, 2019 (both 32 and 64 bit), Office 365, and PowerPoint 2003. It is likely to also work in PowerPoint 2000 and 2007.
  * SVG support is available for Office 365 and recent retail versions of PowerPoint. Support is confirmed for PowerPoint 2021 at least for versions 2108 and above, and likely (although unconfirmed) for PowerPoint 2019 and maybe even PowerPoint 2016 for the same versions. Note that volume licensed versions, which are at version 1808 as of August 2023, do not support SVG conversion to Shape, which is required by IguanaTex.
* LaTeX: [TeXLive](https://www.tug.org/texlive/) or [MiKTeX](http://miktex.org/)
* [GhostScript](http://www.ghostscript.com/download/gsdnld.html) (if the latest version raises issues, try gs9.26)
* [ImageMagick](http://www.imagemagick.org/script/download.php#windows)
* (Optional) [TeX2img](https://github.com/abenori/TeX2img), used for vector graphics output via EMF ([Download](https://www.ms.u-tokyo.ac.jp/~abenori/soft/index.html#TEX2IMG)).

### Mac

* Intel or Apple Silicon Mac
  * On Apple Silicon based Macs, PowerPoint needs to be run with the setting "Open using Rosetta."
* PowerPoint for Mac: 
  * Powerpoint 2019 or Office 365
  * PowerPoint 2016 (Version 16.16.7 190210 or later) (to be confirmed; some features, e.g., SVG, may not work)
* [MacTeX](https://www.tug.org/mactex/): Make sure you install `libgs` for SVG support, by selecting "Customize" at the "Installation Type" step of the MacTex installer. (I haven't found a way to install libgs via TexLive after the initial install, if you know one please let me know)


## Download and Install

### Windows 

1. **Download the .ppam add-in** file from this repository's [Releases page](https://github.com/Jonathan-LeRoux/IguanaTex/releases), or from the [IguanaTex Download page](http://www.jonathanleroux.org/software/iguanatex/download.html), and save it in a [Trusted Location](https://learn.microsoft.com/en-us/DeployOffice/security/trusted-locations) (see [this Microsoft article](https://learn.microsoft.com/en-us/DeployOffice/security/internet-macros-blocked#guidance-on-allowing-vba-macros-to-run-in-files-you-trust)), such as `%appdata%\Microsoft\Addins` (i.e., `C:\Users\user_name\Appdata\Roaming\Microsoft\Addins`).
2. **Load the add-in**: in "File" > "Options" > "Add-Ins" > "Manage:" (lower part of the window), choose "PowerPoint Add-Ins" in the selection box. Then press "Go...", then click  "Add New", select the `.ppam` file in the folder where you downloaded it, then "Close" (if you downloaded the .pptm source and saved it as `.ppam`, it will be in the default Add-In folder).
3. **Create and set a temporary file folder**: IguanaTex needs access to a folder with read/write permissions to store temporary files.
  * The default is "C:\Temp\". If you have write permissions under "C:\", create the folder "C:\Temp\". You're all set.
  * If you cannot create this folder, choose or create a folder with write permission at any other location. In the IguanaTex tab, choose "Main Settings" and put the path to the folder of your choice. You can also use a relative path under the presentation's folder (e.g., ".\" for the presentation folder itself).
4. **Install and set path to GhostScript and ImageMagick**:
  * Set the **full** path to `gswin32c.exe` or `gswin64c.exe` (note the "`c`"!) and to ImageMagick's magick.exe in the "Main Settings" window.
  * Best way to make sure the path is correct is to use the "..." button next to each path and navigate to the correct file.
  * Some default paths include `%USERPROFILE%`. It is recommended to click on "..." to make sure the path gets properly converted to the actual user profile path. 
5. (Optional) **Install and set path to TeX2img**:
  * Only needed for vector graphics support via EMF (compared to SVG, pros: available on all PowerPoint versions, fully modifiable shapes; cons: some displays randomly suffer from distortions)
  * Download from [this link](https://www.ms.u-tokyo.ac.jp/~abenori/soft/index.html#TEX2IMG) (more details on TeX2img on their [Github repo](https://github.com/abenori/TeX2img))
  * After unpacking TeX2img somewhere on your machine, run TeX2img.exe once to let it automatically set the various paths to latex/ghostscript, then set the **full** path to `TeX2imgc.exe` (note the "`c`"!) in the "Main Settings" window.
6. (Optional) **Install LaTeXiT-metadata**:
  * Needed to convert displays generated with [LaTeXiT](https://www.chachatelier.fr/latexit/) on Mac into IguanaTex displays
  * Download [`LaTeXiT-metadata-Win.zip`](https://github.com/Jonathan-LeRoux/IguanaTex/releases/download/v1.60.3/LaTeXiT-metadata-Win.zip) from the Releases page, unzip, and set the path to `LaTeXiT-metadata.exe` in the "Main Settings" window.
  * LaTeXiT-metadata was kindly prepared by Pierre Chatelier, [LaTeXiT](https://www.chachatelier.fr/latexit/)'s author, at my request. Many thanks to him!
  * Source code to be released soon.

**Other settings**:
* If you have a non-standard LaTeX installation, you can specify the folder in which the executables are included.
* If you would like to have the option of using an external editor, e.g., when debugging LaTeX source code, you can specify the path to that editor. If you would like to use that editor by default over the IguanaTex edit window, check the "use as default" checkbox.

### Mac

#### Automatic installation with Homebrew

If you use Homebrew, installation is as simple as:
```bash
brew tap tsung-ju/iguanatexmac
brew install --cask --no-quarantine iguanatexmac latexit-metadata
```
Then follow **5. Verify that paths are set correctly** in the Manual installation instructions below.

For more details (e.g., how to upgrade or uninstall), please see [Tsung-Ju's Homebrew instructions](https://github.com/tsung-ju/homebrew-iguanatexmac).

#### Manual installation

1. **Download the "prebuilt files for Mac" zip** from this repository's [Releases page](https://github.com/Jonathan-LeRoux/IguanaTex/releases)  
There are 3 files to install:
* `IguanaTex.scpt`: AppleScript file for handling file and folder access
* `libIguanaTexHelper.dylib`: library for creating native text views; source code included in the git repo, under "IguanaTexHelper/"
* `IguanaTex_v1_XX_Y.ppam`: main add-in file

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

4. **Load the add-in**: Start PowerPoint (restart if it was running when installing the dylib). From the menu bar, select Tools > PowerPoint Add-ins... > '+' , and choose `IguanaTex_v1_XX_Y.ppam`
  * The first time you click on one of the add-in buttons, you may be notified that `libIguanaTexHelper.dylib` was blocked. Go to the Mac's Settings, then Security and Privacy, and click "Allow Anyway".

5. **Verify that paths are set correctly**: 
  * Click on "Main Settings" in the IguanaTex ribbon tab, and verify that the paths to GhostScript, LaTeX binaries, and libgs.9.dylib (used in SVG conversions) are set correctly by clicking on each "..." button next to them: if the path is correct, this should take you to its location; otherwise, you'll need to navigate to the relevant path. The defaults should match the MacTex installation locations.
  * If you cannot find them, open a terminal and use `locate gs`, `locate pdflatex`, and `locate libgs`.

6. (Optional) **Install LaTeXiT-metadata**:
  * Needed to convert displays generated with [LaTeXiT](https://www.chachatelier.fr/latexit/) on Mac into IguanaTex displays
  * Download [`LaTeXiT-metadata-macos`](https://github.com/Jonathan-LeRoux/IguanaTex/releases/download/v1.60.3/LaTeXiT-metadata-macos) from the Releases page, add executable permission, and either set the path to its location in the "Main Settings" window or copy it to the secure add-in folder:  
  `chmod 755 ./LaTeXiT-metadata-macos`  
  `sudo cp ./LaTeXiT-metadata-macos '/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/'`
  * The first time LaTeXiT-metadata-macos is called by IguanaTex, Mac OS may block it. Go to the Mac's Settings, then Security and Privacy, and click "Allow Anyway".
  * The executable was compiled on Mac OS 10.13 but should work on all versions. Please let me know if you have any issue.
  * LaTeXiT-metadata was kindly prepared by Pierre Chatelier, [LaTeXiT](https://www.chachatelier.fr/latexit/)'s author, at my request. Many thanks to him!
  * [Source code is now public](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-MacOS).

## Non-obvious tricks

IguanaTex's usage should mostly be self-explanatory, but there are a few tips and tricks that may not be.

* Accelerator keys (i.e., keyboard shortcuts): many of IguanaTex's commands ("Generate", "Cancel", etc) can be accessed by using a combination of modifier keys and a single letter. Look for the underlined letter in the corresponding button's text/label.
  * Windows: Alt + letter. For example, instead of clicking on the "<ins>G</ins>enerate" button, you can use `Alt + g`. (This is the standard Office behavior on Windows)
  * Mac: Ctrl + Cmd + letter. For example, instead of clicking on the "<ins>G</ins>enerate" button, you can use `Ctrl + Cmd + g`. (Accelerator keys are not available in the standard Office for Mac, this was specially coded by Tsung-Ju for IguanaTex)

## Known Issues

* "Bitmap" displays created on Mac (which are inserted PDFs) have a small extra margin around them so that they do not appear cropped on Windows. In earlier versions, there was no margin but the display would appear cropped. This seems to be a bug with the way PowerPoint handles some PDFs on Mac, internally storing them as EMF files. The PDFs created by LaTeXiT do not have that issue, however, so there may be a way to circumvent this bug in a future version of IguanaTex.
* IguanaTex macros cannot be added to the Quick Access Toolbar on Mac: this is a [known bug](https://answers.microsoft.com/en-us/msoffice/forum/all/can-add-in-commands-be-added-to-the-quick-access/6872187f-3c17-40ee-8620-80a4068edc82) on which Microsoft is allegedly working.
* There may be some scaling issues when changing the format of a file (Bitmap <-> Vector, or even within the various SVG and EMF Vector formats). The best way to handle this is to use the "Vectorize"/"Rasterize" functions, which regenerate the display in the desired format while fixing the size. One can then further modify the content if needed, and the scaling will be correct.
* For Vector displays, the default "SVG via DVI w/ dvisvgm" is recommended because of issues sometimes observed with other modes: 
  * Some displays obtained via "EMF w/ TeX2img" or "EMF w/ pdfiumdraw" appear distorted. This is a PowerPoint bug that sometimes occurs when ungrouping an EMF file into a Shape object. 
  * Some displays obtained with "SVG via PDF w/ dvisvgm" have symbols or parts of symbol missing. This is because certain lines are represented in PDF by open paths with a certain line width, instead of closed paths, and are thus handled differently by PowerPoint when converting to a Shape object. See [this discussion](https://github.com/mgieseki/dvisvgm/issues/166) for more details.

## License
[![CC BY 3.0][cc-by-image]][cc-by]

This work is licensed under a
[Creative Commons Attribution 3.0 Unported License][cc-by].

[cc-by]: http://creativecommons.org/licenses/by/3.0/
[cc-by-image]: https://i.creativecommons.org/l/by/3.0/88x31.png
[cc-by-shield]: https://img.shields.io/badge/License-CC%20BY%203.0-lightgrey.svg
