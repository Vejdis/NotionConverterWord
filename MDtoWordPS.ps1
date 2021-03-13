Get-ChildItem "DOWNLOADPATH" -file | select -first 1 | Rename-Item -NewName "file.zip"
Move-Item -Path DOWNLOADPATH\file.zip -Destination PREFERABLYEMPTYFOLDER
Expand-Archive -Force PREFERABLYEMPTYFOLDER\File.zip PREFERABLYEMPTYFOLDER
Remove-Item PREFERABLYEMPTYFOLDER\File.zip
Get-ChildItem "PREFERABLYEMPTYFOLDER" -file | select -first 1 | Rename-Item -NewName "File.md"
cd PREFERABLYEMPTYFOLDER
pandoc '.\File.md' -o File.docx
Remove-Item PREFERABLYEMPTYFOLDER\File.md