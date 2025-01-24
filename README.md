# IguanaTex

(C) [Jonathan Le Roux](https://www.jonathanleroux.org/) and Zvika Ben-Haim (Windows), [Tsung-Ju Chiang](https://github.com/tsung-ju) and Jonathan Le Roux (Mac)

IguanaTex is a PowerPoint add-in which allows you to insert LaTeX equations into your PowerPoint presentation on Windows and Mac. It is distributed completely for free, along with its source code.

This repository hosts the source code in a form that can be easily tracked, shared, and discussed (the VBA code is exported using the [ExportVBA macro](https://github.com/Jonathan-LeRoux/IguanaTex/blob/master/ExportVBA.bas)).

The add-in file (.ppam) and its source version (.pptm) can be found in the [Releases](https://github.com/Jonathan-LeRoux/IguanaTex/releases).

## Table of Contents

- [System Requirements](#system-requirements)
  - [Windows](#windows)
  - [Mac](#mac)
- [Download and Install](#download-and-install)
  - [Windows](#windows)
  - [Mac](#mac)
    - [Automatic installation with Homebrew](#automatic-installation-with-homebrew)
    - [Manual installation](#manual-installation)
  - [Other installation settings](#other-installation-settings)
- [Tips, Bugs, and Known Issues](#tips-bugs-and-known-issues)
  - [What to do if something does not work, or does not work as you expected](#what-to-do-if-something-does-not-work-or-does-not-work-as-you-expected)
  - [Debugging an issue](#debugging-an-issue)
  - [Keyboard shortcuts](#keyboard-shortcuts)
  - [Known Issues](#known-issues)
- [Stay up to date: IguanaTex Google Group](#stay-up-to-date-iguanatex-google-group)
- [License](#license)

## System Requirements

### Windows

- OS: Windows 2000 or later (32- or 64-bit).
- PowerPoint:
  - IguanaTex has been tested with Office 365, Office 2019, Office 2021 (including LTSC version), PowerPoint 2003, 2010, 2013, 2016, 2019 (both 32 and 64 bit). It is likely to also work in PowerPoint 2000 and 2007.
  - SVG support is available for Office 365 and recent retail versions of PowerPoint. Support is confirmed for PowerPoint 2021 at least for versions 2108 and above, and likely (although unconfirmed) for PowerPoint 2019 and maybe even PowerPoint 2016 for the same versions. Note that volume licensed versions, which are at version 1808 as of August 2023, do not support SVG conversion to Shape, which is required by IguanaTex.
- LaTeX: [TeXLive](https://www.tug.org/texlive/) or [MiKTeX](http://miktex.org/)
- [GhostScript](http://www.ghostscript.com/download/gsdnld.html) (if the latest version raises issues, try gs9.26)
- [ImageMagick](http://www.imagemagick.org/script/download.php#windows)
- (Optional) [TeX2img](https://github.com/abenori/TeX2img), used for Shape output via EMF ([Download](https://www.ms.u-tokyo.ac.jp/~abenori/soft/index.html#TEX2IMG))
- (Optional) [LaTeXiT-metadata](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-Win), used to convert displays generated with [LaTeXiT](https://www.chachatelier.fr/latexit/) on Mac into IguanaTex displays

### Mac

- Intel or Apple Silicon Mac
  - On Apple Silicon based Macs, PowerPoint needs to be run with the setting "Open using Rosetta."
- PowerPoint for Mac:
  - Office 365, Office 2021 (including LTSC version), Powerpoint 2019, PowerPoint 2016 (Version 16.16.7 190210 or later)
  - SVG support is available for Office 365 and recent retail versions of PowerPoint, including 2019 and 2021. Note that volume licensed (LTSC) versions do not support SVG conversion to Shape, which is required by IguanaTex.
- [XQuartz](https://www.xquartz.org/): Required by `dvisvgm` for SVG support. To be installed before MacTeX's Ghostscript (just rerun the Ghostscript part of MacTeX's installer if you had run it before installing XQuartz)
- [MacTeX](https://www.tug.org/mactex/): For SVG support, make sure to select "Customize" at the "Installation Type" step of the MacTeX installer and check both "Ghostscript Dynamic Library" and "Mutool". [More details about SVG support via `dvisvgm` in MacTeX](https://tug.org/mactex/aboutdvisvgm.html).
- (Optional) [LaTeXiT-metadata](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-MacOS), used to convert [LaTeXiT](https://www.chachatelier.fr/latexit/) displays into IguanaTex displays


## Download and Install

### Windows Installation

1. **Download the .ppam add-in** file from this repository's [Releases page](https://github.com/Jonathan-LeRoux/IguanaTex/releases) and save it in a [Trusted Location](https://learn.microsoft.com/en-us/DeployOffice/security/trusted-locations) (see [this Microsoft article](https://learn.microsoft.com/en-us/DeployOffice/security/internet-macros-blocked#guidance-on-allowing-vba-macros-to-run-in-files-you-trust)), such as `%appdata%\Microsoft\Addins` (i.e., `C:\Users\user_name\Appdata\Roaming\Microsoft\Addins`). If you get a malware warning, try "Trust"-ing the file (Right-Click > Properties). You may have better luck downloading the `.pptm` file, Trusting it, opening it in PowerPoint, and using "Save As" to create your own `.ppam` file.
2. **Load the add-in**: in "File" > "Options" > "Add-Ins" > "Manage:" (lower part of the window), choose "PowerPoint Add-Ins" in the selection box. Then press "Go...", then click  "Add New", select the `.ppam` file in the folder where you downloaded it, then "Close" (if you downloaded the .pptm source and saved it as `.ppam`, it will be in the default Add-In folder).
3. **Create and set a temporary file folder**: IguanaTex needs access to a folder with read/write permissions to store temporary files.
   - The default is "C:\Temp\". If you have write permissions under "C:\", create the folder "C:\Temp\". You're all set.
   - If you cannot create this folder, choose or create a folder with write permission at any other location. In the IguanaTex tab, choose "Main Settings" and put the path to the folder of your choice. You can also use a relative path under the presentation's folder (e.g., ".\" for the presentation folder itself).
4. **Install and set path to GhostScript and ImageMagick** (required for Picture outputs, except if using "Latex (DVI)" engine):
   - Needed to convert intermediate PDF files into PNG before insertion into PowerPoint as a Picture object. In the Picture generation process, all LaTeX engines except "Latex (DVI)" output PDF files. Even when generating a Shape, it is often better to first generate a Picture then convert to Shape: Picture generation is faster and more robust, so it can be convenient to work on a display in Picture mode then do a final conversion to Shape, preserving the size.
   - Set the **full** path to `gswin32c.exe` or `gswin64c.exe` (note the "`c`"!) and to ImageMagick's magick.exe in the "Main Settings" window.
   - Best way to make sure the path is correct is to use the "..." button next to each path and navigate to the correct file.
   - Some default paths include `%USERPROFILE%`. It is recommended to click on "..." to make sure the path gets properly converted to the actual user profile path.
5. (Optional) **Install and set path to TeX2img**:
   - Only needed for vector graphics support via EMF (compared to SVG, pros of EMF are: available on all PowerPoint versions, fully modifiable shapes; cons: some displays randomly suffer from distortions)
   - Download from [this link](https://www.ms.u-tokyo.ac.jp/~abenori/soft/index.html#TEX2IMG) (more details on TeX2img on their [Github repo](https://github.com/abenori/TeX2img))
   - After unpacking TeX2img somewhere on your machine, run TeX2img.exe once to let it automatically set the various paths to latex/ghostscript, then set the **full** path to `TeX2imgc.exe` (note the "`c`"!) in the "Main Settings" window.
6. (Optional) **Install LaTeXiT-metadata**:
   - Needed to convert displays generated with [LaTeXiT](https://www.chachatelier.fr/latexit/) on Mac into IguanaTex displays
   - Download [`LaTeXiT-metadata-Win.zip`](https://github.com/Jonathan-LeRoux/IguanaTex/releases/download/v1.60.3/LaTeXiT-metadata-Win.zip) from the Releases page, unzip, and set the path to `LaTeXiT-metadata.exe` in the "Main Settings" window.
   - LaTeXiT-metadata was kindly prepared by Pierre Chatelier, [LaTeXiT](https://www.chachatelier.fr/latexit/)'s author, at my request. Many thanks to him!
   - [Source code is now public](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-Win).

### Mac Installation

#### Automatic installation with Homebrew

If you use Homebrew, installation is as simple as:

```bash
brew tap tsung-ju/iguanatexmac
brew install --cask --no-quarantine iguanatexmac latexit-metadata
```

Then follow **5. Verify that paths are set correctly** in the Manual installation instructions below.

For more details (e.g., how to **upgrade** or **uninstall**), please see [Tsung-Ju's Homebrew instructions](https://github.com/tsung-ju/homebrew-iguanatexmac).

#### Manual installation

1. **Download the "prebuilt files for Mac" zip** from this repository's [Releases page](https://github.com/Jonathan-LeRoux/IguanaTex/releases)  
There are 3 files to install:
   - `IguanaTex.scpt`: AppleScript file for handling file and folder access
   - `libIguanaTexHelper.dylib`: library for creating native text views; source code included in the git repo, under "IguanaTexHelper/"
   - `IguanaTex_v1_XX_Y.ppam`: main add-in file
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
   - The first time you click on one of the add-in buttons, you may be notified that `libIguanaTexHelper.dylib` was blocked. Go to the Mac's Settings, then Security and Privacy, and click "Allow Anyway".

5. **Verify that paths are set correctly**:
   - Click on "Main Settings" in the IguanaTex ribbon tab, and verify that the following paths are set correctly by clicking on each "..." button next to them. If the path is correct, this should take you to its location; otherwise, you'll need to navigate to the relevant path. The defaults should match the MacTeX installation locations, but your installation may differ.
     - GhostScript
     - LaTeX binaries
     - libgs.dylib (used in SVG conversions; this should only be needed with older versions of MacTeX; leave empty if you get an error, which may happen if you use MacPorts' TeXLive for example)
   - If you cannot find them or if IguanaTex complains that a command did not return, open a terminal and use `locate gs`, `locate pdflatex`, and `locate libgs`.

6. (Optional) **Install LaTeXiT-metadata**:
   - Needed to convert displays generated with [LaTeXiT](https://www.chachatelier.fr/latexit/) on Mac into IguanaTex displays
   - Download [`LaTeXiT-metadata-macos`](https://github.com/Jonathan-LeRoux/IguanaTex/releases/download/v1.60.3/LaTeXiT-metadata-macos) from the Releases page, add executable permission, and either set the path to its location in the "Main Settings" window or copy it to the secure add-in folder:  
  `chmod 755 ./LaTeXiT-metadata-macos`  
  `sudo cp ./LaTeXiT-metadata-macos '/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/'`
   - The first time LaTeXiT-metadata-macos is called by IguanaTex, Mac OS may block it. Go to the Mac's Settings, then Security and Privacy, and click "Allow Anyway".
   - The executable was compiled on Mac OS 10.13 but should work on all versions. Please let me know if you have any issue.
   - LaTeXiT-metadata was kindly prepared by Pierre Chatelier, [LaTeXiT](https://www.chachatelier.fr/latexit/)'s author, at my request. Many thanks to him!
   - [Source code is now public](https://github.com/LaTeXiT-metadata/LaTeXiT-metadata-MacOS).

### Other installation settings

- If you have a non-standard LaTeX installation, you can specify in Main Settings the folder in which the executables are included, or more generally a prefix to be added to all commands (e.g., `wsl -e` for a LaTeX installation under the Windows Subsystem for Linux).
- If you plan to use Tectonic:
  - if you are specifying a path or prefix for the LaTeX installation as explained above, that will be used for the Tectonic executable as well, so please make sure Tectonic is under that path or that it can be called with the specified prefix;
  - if you are not specifying any path or prefix, then the Tectonic executable needs to be on your PATH.
- If you would like to have the option of using an external editor, e.g., when debugging LaTeX source code, you can specify the path to that editor in Main Settings. If you would like to use that editor by default over the IguanaTex edit window, check the "use as default" checkbox.

## Tips, Bugs, and Known Issues

### What to do if something does not work, or does not work as you expected

Most issues originate from some steps of the installation process described above not being followed: please double-check you went through all the steps. A reboot also often helps after a first installation.

If you are having trouble installing or using IguanaTex, please see the [Frequently Asked Questions](https://www.jonathanleroux.org/software/iguanatex/faq.html) and check the [Issues](https://github.com/Jonathan-LeRoux/IguanaTex/issues?q=is%3Aissue) on this repo.

### Debugging an issue

When running into an issue while trying to generate a display, the first thing to do is to check the "Debug" box in the Editor window prior to clicking "Generate". This will step through the process of generating the display, so that we can know where the error occurred, and it will give the option to copy each command so that they can be run in a Terminal or Command Prompt.

If this does not solve the issue, or the issue does not occur during the generation process, the next step is to try to debug in the VBA Editor. To do so:

- open the source `.pptm` file in PowerPoint.
- open the VBA Editor (`Alt+F11` on Windows, `Tools > Macro > Visual Basic Editor` on Mac).
- search for "Macros" under "Module" in the exploration pane on the left.
- place a breakpoint, for example at Line 7 (`Load LatexForm` under `NewLatexEquation()`) by clicking in the margin.
- Launch the display generation process:
  - on Windows, click on the "New LaTeX Display" button in the IguanaTex ribbon (if the add-in is loaded, you will likely have two IguanaTex tabs in the ribbon, one for the loaded add-in and the other for the `.pptm` file: just try one, and if the IguanaTex window appears without hitting the breakpoint, try the other), or in the VBA Editor click `Tools >  Macros...` and select `NewLatexEquation` and `Run`.
  - on Mac, clicking on the buttons in the ribbon does not work for `.pptm` files, so instead click `Tools > Macro > Macros...` and select `NewLatexEquation` and `Run`.
- The code will stop at the breakpoint.
- Step Over (Shift+F8 on Windows, Shift+⌘+O on Mac) until hitting the bug. If the bug occurs on a line calling another function, you can run again and then Step Into (F8 on Windows, Shift+⌘+I on Mac) when you reach that line.
- Eventually, you'll reach the actual line causing the bug. Now, either try to fix it, or open an issue.

### Keyboard shortcuts

Accelerator keys (i.e., keyboard shortcuts): many of IguanaTex's commands ("Generate", "Cancel", etc) can be accessed by using a combination of modifier keys and a single letter. Look for the underlined letter in the corresponding button's text/label.

- Windows: Alt + letter. For example, instead of clicking on the "<ins>G</ins>enerate" button, you can use `Alt + g`. (This is the standard Office behavior on Windows)
- Mac: Ctrl + Cmd + letter. For example, instead of clicking on the "<ins>G</ins>enerate" button, you can use `Ctrl + ⌘ + g`. (Accelerator keys are not available in the standard Office for Mac, this was specially coded by Tsung-Ju for IguanaTex)

### Known Issues

- "Picture" displays created on Mac (which are inserted PDFs) appear cropped on Windows ([Issue #32](https://github.com/Jonathan-LeRoux/IguanaTex/issues/32)). Regenerating them on Windows fixes the issue. This seems to be a bug with the way PowerPoint handles some PDFs on Mac, internally storing them as EMF files. The PDFs created by LaTeXiT do not have that issue, however, so there may be a way to circumvent this bug in a future version of IguanaTex.
- IguanaTex macros cannot be added to the Quick Access Toolbar on Mac ([Issue #23](https://github.com/Jonathan-LeRoux/IguanaTex/issues/23)): this is a [known bug](https://answers.microsoft.com/en-us/msoffice/forum/all/can-add-in-commands-be-added-to-the-quick-access/6872187f-3c17-40ee-8620-80a4068edc82) on which Microsoft is allegedly working, although there has been no progress for multiple years.
- There may be some scaling issues when changing the format of a file (Picture <-> Shape, or even within the various SVG and EMF Shape formats). The best way to handle this is to use the "Convert to Shape"/"Convert to Picture" functions, which regenerate the display in the desired format while keeping the size fixed. One can then further modify the content if needed, and the scaling will be correct.
- For Shape (i.e., vector graphics) displays, the default "SVG via DVI w/ dvisvgm" is recommended because of issues sometimes observed with other modes:
  - Some displays obtained via "EMF w/ TeX2img" or "EMF w/ pdfiumdraw" appear distorted. This is a PowerPoint bug that sometimes occurs when ungrouping an EMF file into a Shape object.
  - Some displays obtained with "SVG via PDF w/ dvisvgm" have symbols or parts of symbol missing. This is because certain lines are represented in PDF by open paths with a certain line width, instead of closed paths, and are thus handled differently by PowerPoint when converting to a Shape object. See [this discussion](https://github.com/mgieseki/dvisvgm/issues/166) for more details.

## Stay up to date: IguanaTex Google Group

To be informed of the release of new versions, you can subscribe to the [IguanaTex Google Group](https://groups.google.com/d/forum/iguanatex).

## License

[![CC BY 3.0][cc-by-image]][cc-by]

This work is licensed under a
[Creative Commons Attribution 3.0 Unported License][cc-by].

[cc-by]: http://creativecommons.org/licenses/by/3.0/
[cc-by-image]: https://i.creativecommons.org/l/by/3.0/88x31.png
[cc-by-shield]: https://img.shields.io/badge/License-CC%20BY%203.0-lightgrey.svg
