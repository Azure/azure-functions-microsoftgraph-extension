# Upgrade instructions
1. Stop your function application
2. Go to Platform features > Advanced tools (Kudu)
3. On Kudu, go to the Debug Console
4. Navigate to the directory D:/home/site/wwwroot
5. Put the file extensions.csproj into this directory
6. Run the command: dotnet build -o bin extensions.csproj
7. Restart your function application.



