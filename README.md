# Windows Mass Uninstall

Sometimes you need to uninstall a lot of applications. 
For example, installing any Microsoft Visual Studio, even in minimal configuration, 
leads to installation to numerous add-ons and applications, which were not explicitly selected for installation.
Selecting only C++ compiler for installation, you will definitely receive Microsoft SQL Server Compact (Developer) Edition.
This applicable to any software distributions, consist of multiple packages.

This script releases you of this monotonous work. For your convenience, you can read a list of all installed application, to know their actual names:

`win_mass_uninstall.py --installed-apps > applications.txt`

After that, pick up only applications you want to uninstall, and for safety rename the file:

`win_mass_uninstall.py --uninstall > uninstall_list.txt`

WMIC is used for actual uninstalling, and all requirements applicable for WMIC project to the Mass Uninstall script, 
like you need to run it under the Administrator account.
