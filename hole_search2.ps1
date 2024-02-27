
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Xml

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$dialogResult = $folderBrowser.ShowDialog()

if ($dialogResult -eq "OK") {
    $folderPath = $folderBrowser.SelectedPath
    $files = Get-ChildItem -Path $folderPath -Filter "*.xml"

    $totalCount = 0
    $count35 = 0

    foreach ($file in $files) {
        $xmlContent = [System.Xml.XmlDocument](Get-Content -Path $file.FullName)
        $cadNodes = $xmlContent.SelectNodes('//CAD')

        foreach ($cadNode in $cadNodes) {
            $diameterNodes = $cadNode.SelectNodes('.//Diameter')
            $totalCount += $diameterNodes.Count

            foreach ($diameterNode in $diameterNodes) {
                if ([float]$diameterNode.'#text' -eq 35.00) {
                    $typeNoNode = $cadNode.SelectSingleNode('.//TypeNo')

                    if ($typeNoNode -and $typeNoNode.'#text' -eq "1") {
                        $count35 += 1
                    }
                }
            }
        }
    }

    Write-Host "Общее количество отверстий: $totalCount"
    Write-Host "Петли: $count35"
    Write-Host "Отдельно отверстия: $($totalCount - $count35)"
}
pause