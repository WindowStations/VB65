# Upgrade Visual Basic 6.0 IDE to Visual Basic 6.5 IDE with the VBA 6.5 SDK. - Run As Administrator - Still in BETA
VB65 Requirements:
1. It is required to have a licensed copy of Visual Basic 6.0 and the VBA SDK is installed on the development machine.  The SDK zip files above are stored in WinRAR format.  They were split up only because of the upload file size limitations of GitHub.
2. Install the VBA SDK 6.0 version 6.5. 
3. Download VB65.exe to the desired location.
4. Apply compatibility settings to VB65.exe for "Windows XP (Service pack 2)".
5. Always "Run as Administrator".
6. Visual Basic for Applications is an embeddable BASIC language software product comprised of the following components:
* VBA
* APC
* Microsoft Forms
* Core Technology   
* End User Documentation and derivatives
* VBA Core Installer Package

VBA 
These are the binary files that make up the core VBA deliverable.  This includes the language runtime and user interface components including editing, debugging, project management, property control. It also includes the files need for Multi-threading and the multi-threading runtime. 
VBE6.DLL, VBE6EXT.OLB, SCP32.DLL, VBE6INTL.DLL, VBAME.DLL, LINK.EXE, MSPDB60.DLL, MTDSR.DLL, VBA6MTRT.DLL, VB6DEBUG.DLL 
  
APC 
These are the binary files that make up the application programming interface (API) layer to integrate VBA.
APC65.DLL, APC60ITL.DLL 

Microsoft Forms
This provides a complete visual editing and dialog design environment. 
FM20.DLL, FM20ENU.DLL, RICHED20.DLL

Core Technology
VBA is dependent on a set of MICROSOFT core proprietary technology.  These technologies are to be consumed by VBA only.  Hosts of VBA are forbidden from accessing core technology directly. 
MSO.DLL, MSOINTL.DLL, SELFCERT.EXE, SIGNER.DLL, MSVBVM60.DLL, MSSTDFMT.DLL, MSSTKPRP.DLL, HLP95EN.DLL    

VBA Core Installer Package (MSI)
MICROSOFT provides a setup package that installs all the core VBA technologies onto an end users machine.  This prevents accidental disabling of VBA functionality due to an improper installation of VBA components. 

