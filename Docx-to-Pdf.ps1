$Files = Get-ChildItem 'C:\Temp\converttopdftest\*.Docx'
$Word = New-Object -ComObject Word.Application 

Foreach ($File in $Files) {
    
    $Doc = $Word.Documents.Open($File.fullname)

    $Name = ($Doc.Fullname).replace("docx","pdf")

    $Doc.saveas([ref] $Name, [ref] 17)

    $Doc.close()

}


