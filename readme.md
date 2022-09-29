


mkdir cli2
cd cli2
code .
dotnet new sln

dotnet new classlib -o docx_lib
cd docx_lib
dotnet add package Markdig --version 0.30.4
dotnet add package NS.HtmlToOpenXml --version 1.1.0
cd ..

dotnet new console -o docx_md
dotnet new blazorwasm -o docx_web

dotnet sln add docx_md/docx_md.csproj
dotnet sln add docx_web/docx_web.csproj 
dotnet sln add docx_lib/docx_lib.csproj 

dotnet add docx_md/docx_md.csproj reference docx_lib/docx_lib.csproj

dotnet add docx_web/docx_web.csproj reference docx_lib/docx_lib.csproj

dotnet build
dotnet publish -c release
cd bin/Release/net6.0/publish/wwwroot
surge . mddocx2.surge.sh
