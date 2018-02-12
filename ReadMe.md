## Requirements
- (First time, or any time you wish to upgrade USMT) Microsoft's User State Migration Tool - This is available by Installing the Windows Assessment and Deployment Toolkit. Search MS to find it. USMT is a tickable option when installing the ADK - To my knowledge, there is not a specific USMT version requirement.

## Usage Instructions
1. On the computer where the Windows ADK is installed: copy the architecture specific folders (x86, amd64, arm64) from "$env:ProgramFiles(x86)\Windows Kits\$ver\Assessment and Deployment Kit\User State Migration Tool\" to the root MigAssistant Folder
1. Customize MigAssistant.exe.config to suit your needs
1. (Optional) Copy MigApp.XML and/or MigUser.XML from any of the architecture-specific folders to the top-level MigAssistant Folder
    1. Customize MigApp.XML and/or MigUser.XML
    1. Copy the prepared MigAssistant Folder to a network location (Recommended) or copy it locally to a machine you wish to migrate.
1. Run MigAssistant.exe as Administrator

## Build/Debug Instructions
1. Builds without modification to AnyCPU
1. Same requirements as above to Debug
    1. Follow the Usage instructions regarding installing and copying USMT, but your destination folder is the Debug output folder for the project in Visual Studio.
