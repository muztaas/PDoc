# Bundle Python first
.\bundle-python.ps1

# Build the solution
dotnet restore ..\PDoc.sln
dotnet build ..\PDoc.sln -c Release
dotnet publish ..\src\PDoc\PDoc.csproj -c Release -r win-x64 --self-contained

# Create installer
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss