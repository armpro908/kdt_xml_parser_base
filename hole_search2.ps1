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
    $totalLineLength = 0

    foreach ($cadNode in $cadNodes) {
        $totalCount, $count35, $totalLineLength = ProcessCadNode -CadNode $cadNode -Quantity $quantity -TotalCount $totalCount -Count35 $count35 -TotalLineLength $totalLineLength

        if ($count35 -gt 0) {
            $foundHingeInFile = $true
        }
    }

    return [PSCustomObject]@{
        TotalCount = $totalCount
        Count35 = $count35
        TotalLineLength = $totalLineLength
        FoundHingeInFile = $foundHingeInFile
        FileName = $file.Name
    }
}


function ProcessCadNode($CadNode, $Quantity, $TotalCount, $Count35, $TotalLineLength) {
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

    $typeNoNode = $CadNode.SelectSingleNode('.//TypeNo')
    if ($typeNoNode -and $typeNoNode.'#text' -eq "3") {
        $beginX = [int]$CadNode.SelectSingleNode('.//BeginX').'#text'
        $beginY = [int]$CadNode.SelectSingleNode('.//BeginY').'#text'
        $endX = [int]$CadNode.SelectSingleNode('.//EndX').'#text'
        $endY = [int]$CadNode.SelectSingleNode('.//EndY').'#text'

        if ($beginX -eq $endX) {
            $line_length = [Math]::Abs($endY - $beginY)
        } elseif ($beginY -eq $endY) {
            $line_length = [Math]::Abs($endX - $beginX)
        }

        $TotalLineLength += $line_length * $Quantity
    }

    return $TotalCount, $Count35, $TotalLineLength
}

$folderPath = SelectFolder
$files = GetXmlFiles $folderPath

$results = @()

foreach ($file in $files) {
    $results += ProcessXmlFile $file
}

$totalCount = ($results | Measure-Object -Property TotalCount -Sum).Sum
$count35 = ($results | Measure-Object -Property Count35 -Sum).Sum
$totalLineLength = ($results | Measure-Object -Property TotalLineLength -Sum).Sum
$filesWithHinges = ($results | Where-Object { $_.FoundHingeInFile }).FileName


Write-Host "==============================================================="
Write-Host "                         Отчет                                 "
Write-Host "==============================================================="

Write-Host "Общее количество отверстий: " -NoNewline
Write-Host $totalCount -ForegroundColor Cyan

Write-Host "Петли: " -NoNewline
Write-Host $count35 -ForegroundColor Cyan

Write-Host "Файлы с петлями: " -NoNewline
Write-Host $($filesWithHinges -join ', ') -ForegroundColor Cyan

Write-Host "Отдельно отверстия: " -NoNewline
Write-Host $($totalCount - $count35) -ForegroundColor Cyan

Write-Host "Суммарная длина пазов (в метрах): " -NoNewline
Write-Host $([Math]::Round($totalLineLength / 1000, 2)) -ForegroundColor Cyan

Write-Host "==============================================================="


pause
