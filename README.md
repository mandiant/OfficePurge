# OfficePurge

VBA purge your Office documents with OfficePurge. VBA purging removes P-code from module streams within Office documents. Documents that only contain source code and no compiled code are more likely to evade AV detection and YARA rules. 
Read more <a href="https://blog.nviso.eu/2020/02/25/evidence-of-vba-purging-found-in-malicious-documents/">here</a>.

OfficePurge supports VBA purging Microsoft Office Word (.doc), Excel (.xls), and Publisher (.pub) documents. Original and purged documents for each supported file type with a macro that will spawn calc.exe can be found in `sample-data` folder.

Author: Andrew Oliveau (@AndrewOliveau)


# INSTALLATION/BUILDING

## Pre-Compiled 

* Use the pre-compiled binary in the <a href="https://github.com/fireeye/OfficePurge/releases">Releases</a> section

## Building Yourself

Take the below steps to setup Visual Studio in order to compile the project yourself. This requires a couple of .NET libraries that can be installed from the NuGet package manager.

### Libraries Used
The below 3rd party libraries are used in this project.

| Library | URL | License |
| ------------- | ------------- | ------------- |
| OpenMCDF  | [https://github.com/ironfede/openmcdf](https://github.com/ironfede/openmcdf) |  MPL-2.0 License   |
| Fody  | [https://github.com/Fody/Fody](https://github.com/Fody/Fody) | MIT License  |
| Kavod.Vba.Compression | [https://github.com/rossknudsen/Kavod.Vba.Compression](https://github.com/rossknudsen/Kavod.Vba.Compression) |  MIT License |

### Steps to Build
* This project requires .NET framework 4.7
* Load the Visual Studio project up and go to "Tools" --> "NuGet Package Manager" --> "Package Manager Settings"
* Go to "NuGet Package Manager" --> "Package Sources"
* Add a package source with the URL "https://api.nuget.org/v3/index.json"
* Install the Costura.Fody NuGet package. The older version of Costura.Fody (3.3.3) is needed, so that you do not need Visual Studio 2019.
  * `Install-Package Costura.Fody -Version 3.3.3`
* Install OpenMcdf to manipulate Microsoft Compound Document Files. OpenMcdf version (2.2.1.3) is needed so that the current code works correctly.
  * `Install-Package OpenMcdf -Version 2.2.1.3`
* Install Fody
  * `Install-Package Fody -Version 4.0.2` 
* You can now modify and build the project yourself!


# ARGUMENTS/OPTIONS
* <b>-d </b> - Document type to VBA purge (word, excel, publisher) 
* <b>-f </b> - Document filename to VBA purge 
* <b>-m </b> - Module within document to VBA purge (ex. Module1)
* <b>-l </b> - List modules in a document 
* <b>-h </b> - Show help menu

# EXAMPLES

* `OfficePurge.exe -d word -f .\malicious.doc -m NewMacros`
* `OfficePurge.exe -d excel -f .\payroll.xls -m Module1`
* `OfficePurge.exe -d publisher -f .\donuts.pub -m ThisDocument`
* `OfficePurge.exe -d word -f .\malicious.doc -l`

# REFERENCES
* Didier Steven's VBA purging article <a href="https://blog.nviso.eu/2020/02/25/evidence-of-vba-purging-found-in-malicious-documents/">here</a>
* EvilClippy for parts of code <a href="https://github.com/outflanknl/EvilClippy">here</a>
