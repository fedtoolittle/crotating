// Serious Version

Crotating


Crotating is a high-performance Windows-based utility designed to automate the ingestion and analysis of work timecard data. Originally developed on legacy .NET Framework, this version has been fully modernized to .NET 10, leveraging SDK-style project architecture and high-efficiency Excel processing via EPPlus.

🚀 Features
Asynchronous Excel Ingestion: Rapidly parses large-scale timecard files using the EPPlus engine.

Intelligent Logic: Automatically identifies `TOTAL` rows and handles missing data points without crashing (hopefully).

Modern .NET 10 Engine: Utilizes the latest C# features, including Nullable Reference Types for increased runtime stability.

Self-Contained Deployment: Can be compiled into a single executable that runs on Windows without requiring a pre-installed .NET runtime.

🛠 Technical Stack
Language: C# 14.0

Framework: .NET 10.0 (WinForms)

Excel Engine: EPPlus 7+

Pattern: Service-Oriented Architecture (IWorkEntryReader)

📦 Installation & Setup
Prerequisites
Visual Studio 2026 (or VS Code with the C# Dev Kit)

.NET 10.0 SDK

Building from Source

Clone the repository:
```
git clone https://github.com/fedtoolittle/crotating.git
```
Navigate to the project directory:
```
cd crotating
```
Restore dependencies and build:
```
dotnet build -c Release
```
🏗 Deployment
To generate a standalone, single-file executable for distribution:
```
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish
```
📖 Usage
Launch Crotating.exe.

Select your source timecard Excel file.

The system will validate the data structure and generate a summarized work report.

Export the results to a standardized summary file for payroll or project management.

⚖️ License
This project is licensed under the MIT License - see the `LICENSE` file (pending) for details. Note that the EPPlus library requires a separate license context for commercial use.


// More Serious Version

This is a useless program that will never be used by anyone else because it has zero real world application 99.99% of the time

All it does is take an excel document in a very specific format and converts it into another very specific format

Literally (x, y) -> (y, x)

To use it simply select the file that Crab al vomited out and select the correct file type. Press run and if nothing breaks hit export and voila.
If anything breaks then only god knows what went wrong because i can't be bothered to write error messages.

Also it has to be opened from Visual Studio because i am not a developer and screw passing around installer.
