

```
mkdir cli2
cd cli2
code .
dotnet new sln

dotnet new classlib -o docx_lib
cd docx_lib
dotnet add package Markdig --version 0.30.4
dotnet add package NS.HtmlToOpenXml --version 1.1.0
cd ..

dotnet new console -o office_convert
dotnet new blazorwasm -o docx_web
dotnet new mstest -o convert_test
dotnet new classlib -o docx_lib

dotnet sln add office_convert/office_convert.csproj
dotnet sln add docx_web/docx_web.csproj 
dotnet sln add docx_lib/docx_lib.csproj 
dotnet sln add convert_test/convert_test.csproj 

dotnet add office_convert/office_convert.csproj reference docx_lib/docx_lib.csproj
dotnet add convert_test/convert_test.csproj reference office_convert/office_convert.csproj

dotnet add docx_web/docx_web.csproj reference docx_lib/docx_lib.csproj

dotnet build
dotnet publish -c release
cd bin/Release/net6.0/publish/wwwroot
surge . mddocx2.surge.sh

```


https://github.com/microsoft/vstest/issues/799
using (var writer = new System.IO.StreamWriter(System.Console.OpenStandardOutput()))
       writer.WriteLine("This will show up!");

dotnet test -l "console;verbosity=detailed" 

dotnet mstest
https://executecommands.com/dotnet-core-mstest-project/

Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
https://stackoverflow.com/questions/10204091/how-to-get-directory-while-running-unit-test
