Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Xml

function SelectFolder {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialogResult = $folderBrowser.ShowDialog()

    if ($dialogResult -eq "OK") {
        return $folderBrowser.SelectedPath
    } else {
        Write-Host "Не выбран каталог"
        exit
    }
}

function GetXmlFiles($folderPath) {
    return Get-ChildItem -Path $folderPath -Filter "*.xml"
}

function ProcessXmlFile($file) {
    $xmlContent = [System.Xml.XmlDocument](Get-Content -Path $file.FullName)
    $cadNodes = $xmlContent.SelectNodes('//CAD')

    $quantity = 1
    if ($file.Name -match '--(\d+)\.XML$') {
        $quantity = [int]$matches[1]
    }

    $foundHingeInFile = $false
    $totalCount = 0
    $count35 = 0

    foreach ($cadNode in $cadNodes) {
        $totalCount, $count35 = ProcessCadNode -CadNode $cadNode -Quantity $quantity -TotalCount $totalCount -Count35 $count35

        if ($count35 -gt 0) {
            $foundHingeInFile = $true
        }
    }

    return [PSCustomObject]@{
        TotalCount = $totalCount
        Count35 = $count35
        FoundHingeInFile = $foundHingeInFile
        FileName = $file.Name
    }
}

function ProcessCadNode($CadNode, $Quantity, $TotalCount, $Count35) {
    $diameterNodes = $CadNode.SelectNodes('.//Diameter')
    $TotalCount += $diameterNodes.Count * $Quantity

    foreach ($diameterNode in $diameterNodes) {
        if ([float]$diameterNode.'#text' -eq 35.00) {
            $typeNoNode = $CadNode.SelectSingleNode('.//TypeNo')

            if ($typeNoNode -and $typeNoNode.'#text' -eq "1") {
                $Count35 += 1 * $Quantity
            }
        }
    }

    return $TotalCount, $Count35
}

$folderPath = SelectFolder
$files = GetXmlFiles $folderPath

$results = @()

foreach ($file in $files) {
    $results += ProcessXmlFile $file
}

$totalCount = ($results | Measure-Object -Property TotalCount -Sum).Sum
$count35 = ($results | Measure-Object -Property Count35 -Sum).Sum
$filesWithHinges = ($results | Where-Object { $_.FoundHingeInFile }).FileName

Write-Host "==============================================================="
Write-Host "                         Отчет                                 "
Write-Host "==============================================================="

Write-Host "Общее количество отверстий: " -NoNewline
Write-Host $totalCount -ForegroundColor DarkRed

Write-Host "Петли: " -NoNewline
Write-Host $count35 -ForegroundColor DarkRed

Write-Host "Отдельно отверстия: " -NoNewline
Write-Host $($totalCount - $count35) -ForegroundColor DarkRed

Write-Host "Файлы с петлями: " -NoNewline
Write-Host $($filesWithHinges -join ', ') -ForegroundColor DarkRed

Write-Host "==============================================================="

pause

