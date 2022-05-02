# TwinBasicSevenZip
7-Zip COM library compatible for use in VBA (32-bit and 64-bit), VB6 and twinBASIC

This is a proof of concept to demonstrate the capabilities of twinBASIC both as a COM consumer and producer. Traditionally, VB6/VBA solutions can only work with [Automation-compatible](https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/bbb05720-f724-45c7-8d17-f83c3d1a3961) libraries. To work with non-Automation COM libraries requires use of additional tools or hacks to work around the limitations. 7-Zip DLL was chosen to demonstrate that twinBASIC is capable of natively authoring a solution that can take the COM-like interfaces from 7-Zip's DLLs and work with it and then expose it as an Automation-compatible object or even as a simple `Declare` statement. 

This is a very rough and there is much to tap. The error handling is minimal and the testing coverage need to be better. 

# Sample Usage

![image](https://user-images.githubusercontent.com/2367644/166168441-e3cd63a4-2f2e-4e89-a3a2-7d97196a8544.png)

There are two ways to use the DLL:

1. As a `Declare` statement:

This bypass the need to register the DLL and use the COM objects, allowing you to directly extract a supported archive format into a directory of your choosing. The sample VBA/VB6 compatible code follows:

```
Private Declare PtrSafe Sub Extract Lib "C:\Temp\SevenZipArchive_win64.dll" (ArchivePath As LongPtr, DestinationFolder As LongPtr)

Public Sub DemoDeclare()
    Extract StrPtr("C:\Temp\Test.7z"), StrPtr("C:\Temp\Test7z_Extract")
End Sub
```

This requires placing one of 7-Zip's DLL in the same folder as the `SevenZipArchive_winXX.dll` in order for the twinBASIC DLL to work. You can use `7z.dll`, `7za.dll` or `7zxa.dll` depending on your requirements. For example, if you want to unzip `.7z` file, then any those DLL will work. However, if you want to unzip using `.zip` files, you must use `7z.dll`. 

2. As a COM objects:

You can register the library using `regsvr32.exe` (keeping in mind to use the appropriate `regsvr32.exe` to match the bitness) and then reference the library from your VBx project.

![image](https://user-images.githubusercontent.com/2367644/166168701-444c7125-2389-4e8b-910a-21f83a3f8d06.png)

The `ArchiveFactory` allow you to specify where to locate the 7-Zip's DLL files (which can be one of `7z.dll`, `7za.dll` or `7zxa.dll`) If it's not specified explicitly, it defaults to the same folder as the `SevenZipArchive_winXX.dll` preferring to load the DLL in the given order. You can then create an `ArchiveExtracotr` from the factory as shown:

```
Public Sub DemoCom()
    SevenZipArchive.ArchiveFactory.ArchiveLibPath = "C:\Program Files\7-Zip\7z.dll"
    
    Dim ae As SevenZipArchive.ArchiveExtractor
    Set ae = SevenZipArchive.ArchiveFactory.CreateArchiveExtractor("C:\Temp\Test.7z")
    
    ae.Extract "C:\Temp\Test7z_Com"
End Sub
```

# TODOs:

As this is in alpha, there are few numbers of features that are possible but not implemented:

* Implement the compress routine
* Handle subfolders within archive files and preserving various file attributes (e.g. timestamps)
* Provide a progress dialog 
* Provide a event to allow the consuming VBx/tB code to respond to each file being extracted/compressed
* Provide more control over the format supported (currently it depends on the file extension but there are other ways such as looking for an unique signature within the file)
* Update the keywords that are planned to be replaced (see: https://github.com/WaynePhillipsEA/twinbasic/issues/806 )
* Tests, tests, tests!

# Contributing & Building

To build, you must have twinBASIC compiler. You can obtain the latest from here:
https://github.com/WaynePhillipsEA/twinbasic/issues/772

You can then import the source code contained in the `Twinbasic` folder using command like this:
```
...\bin\twinBASIC_win64.exe" import .\SevenZipArchive.twinproj .\Source --overwrite
```

For more details, refer to [the wiki](https://github.com/bclothier/TwinBasicSevenZip/wiki).
