
dotnet build
dotnet nuget push -s https://nuget.dtvl.com.tw -k %NUGET_API_KEY% bin/Debug/Rugal.OpenReport.1.0.5.14.nupkg

pause