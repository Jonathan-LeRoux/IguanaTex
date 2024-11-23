Attribute VB_Name = "Defaults"
Option Explicit

#If Mac Then
Public Const DEFAULT_TEMP_DIR As String = vbNullString
Public Const DEFAULT_TEX_EXE_PATH As String = "/Library/TeX/texbin/"
Public Const DEFAULT_TEX_EXTRA_PATH As String = vbNullString
Public Const DEFAULT_LIBGS As String = "/opt/local/lib/libgs.9.dylib"
Public Const DEFAULT_VECTOR_OUTPUT_TYPE As String = "dvisvgm"
Public Const DEFAULT_PICTURE_OUTPUT_TYPE As String = "PDF"
Public Const DEFAULT_GS_COMMAND As String = "/usr/local/bin/gs"
Public Const DEFAULT_IM_CONV As String = "/usr/local/bin/convert"
Public Const DEFAULT_TEX2IMG_COMMAND As String = "/usr/local/bin/tex2img"
Public Const DEFAULT_EDITOR As String = "open -b 'texstudio'"
Public Const DEFAULT_ADDIN_FOLDER As String = "/Library/Application Support/Microsoft/Office365/User Content.localized/Add-Ins.localized/"
Public Const DEFAULT_LATEXIT_METADATA_COMMAND As String = DEFAULT_ADDIN_FOLDER & "LaTeXiT-metadata-macos"
Public Const NEWLINE As String = vbLf
Public Const PathSep As String = "/"
Public Const WrongPathSep As String = "\"

#Else
Public Const DEFAULT_TEMP_DIR As String = "c:\temp\"
Public Const DEFAULT_TEX_EXE_PATH As String = vbNullString
Public Const DEFAULT_TEX_EXTRA_PATH As String = vbNullString
Public Const DEFAULT_LIBGS As String = vbNullString
Public Const DEFAULT_VECTOR_OUTPUT_TYPE As String = "dvisvgm"
Public Const DEFAULT_PICTURE_OUTPUT_TYPE As String = "PNG"
Public Const DEFAULT_GS_COMMAND As String = "C:\Program Files (x86)\gs\gs9.15\bin\gswin32c.exe"
Public Const DEFAULT_IM_CONV As String = "C:\Program Files\ImageMagick\magick.exe"
Public Const DEFAULT_TEX2IMG_COMMAND As String = "%USERPROFILE%\Downloads\TeX2img\TeX2imgc.exe"
Public Const DEFAULT_EDITOR As String = "C:\Program Files (x86)\TeXstudio\texstudio.exe"
Public Const DEFAULT_LATEXIT_METADATA_COMMAND As String = "%USERPROFILE%\Downloads\LaTeXiT-metadata\LaTeXiT-metadata-win.exe"
Public Const NEWLINE As String = vbCrLf
Public Const PathSep As String = "\"
Public Const WrongPathSep As String = "/"

#End If

Public Const IGUANATEX_VERSION As Integer = 162

Public Const DEFAULT_LATEX_CODE As String = "\documentclass{article}" & NEWLINE & "\usepackage{amsmath}" & NEWLINE & "\pagestyle{empty}" & NEWLINE & _
                                            "\begin{document}" & NEWLINE & NEWLINE & NEWLINE & NEWLINE & NEWLINE & "\end{document}"
Public Const DEFAULT_LATEX_CODE_PRE As String = "\documentclass{article}" & NEWLINE & "\usepackage{amsmath}" & NEWLINE & "\pagestyle{empty}" & NEWLINE & _
                                                "\begin{document}" & NEWLINE & NEWLINE
Public Const DEFAULT_LATEX_CODE_POST As String = NEWLINE & NEWLINE & "\end{document}"
