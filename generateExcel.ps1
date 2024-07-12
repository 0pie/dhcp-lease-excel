try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Error "Erreur lors de l'importation du module ImportExcel"
    exit
}

$dhcpServer = "VM-DC01"
$scopeId = "192.168.50.0"

try {
    $dhcpLeases = Get-DhcpServerv4Lease -ComputerName $dhcpServer -ScopeId $scopeId -ErrorAction Stop
} catch {
    Write-Error "Erreur lors de l'échange avec $dhcpServer"
    exit
}

$data = @()

for ($i = 1; $i -le 254; $i++) {
    $ip = "192.168.50.$i"

    $lease = $dhcpLeases | Where-Object { $_.IPAddress -eq $ip }

    if ($lease) {
        $status = switch ($lease.AddressState) {
            "Active"    { "Utilisée" }
            "Inactive"  { "Libre" }
            default     { "Réservée" }
        }
    } else {
        $status = "Libre"
    }

    $data += [PSCustomObject]@{
        "Adresse IP"   = $ip
        "Masque"       = "255.255.255.0"
        "Statut"       = $status
        "Nom"          = if ($lease) { $lease.HostName } else { "" }
        "Adresse MAC"  = if ($lease) { $lease.ClientId } else { "" }
    }
}

$excelFilePath = "\\192.168.50.227\Public\INFORMATIQUE\InventaireMachine.xlsx"
$excelPackage = Open-ExcelPackage -Path $excelFilePath

$worksheet = $excelPackage.Workbook.Worksheets["VLAN1"]
if (-not $worksheet) {
    $worksheet = $excelPackage.Workbook.Worksheets.Add("VLAN1")
}

$data | Export-Excel -Worksheet $worksheet -TableName "DHCP" -AutoSize -ClearSheet # ouverture

$rangeHeader = $worksheet.Cells["A1:G1"]
$rangeHeader.Style.Font.Color.SetColor([System.Drawing.Color]::White)
$rangeHeader.Style.Fill.PatternType = 'Solid'
$rangeHeader.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Blue)

$rowIndex = 2
foreach ($row in $worksheet.Cells["A2:G" + $worksheet.Dimension.End.Row]) {
    if ($rowIndex % 2 -eq 0) { 
        $row.Style.Fill.PatternType = 'Solid'
        $row.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightBlue)
    }
    $row.Style.Border.Top.Style = 'Thin'
    $row.Style.Border.Left.Style = 'Thin'
    $row.Style.Border.Right.Style = 'Thin'
    $row.Style.Border.Bottom.Style = 'Thin'
    $row.Style.Border.Top.Color.SetColor([System.Drawing.Color]::DarkBlue)
    $row.Style.Border.Left.Color.SetColor([System.Drawing.Color]::DarkBlue)
    $row.Style.Border.Right.Color.SetColor([System.Drawing.Color]::DarkBlue)
    $row.Style.Border.Bottom.Color.SetColor([System.Drawing.Color]::DarkBlue)
    $rowIndex++
}

$worksheet2 = $excelPackage.Workbook.Worksheets["Switch"]
if (-not $worksheet2) {
    $worksheet = $excelPackage.Workbook.Worksheets.Add("Switch")
}


$sourceFilePath = "\\192.168.50.227\Public\INFORMATIQUE\InventaireEquipement.xlsx"
$sourceWork = Open-ExcelPackage -Path $sourceFilePath

foreach ($sourceWorksheet in $sourceWork.Workbook.Worksheets) {
    $destinationWorksheet = $excelPackage.Workbook.Worksheets.Add($sourceWorksheet.Name)
    $destinationWorksheet.Cells[$sourceWorksheet.Dimension.Address].Value = $sourceWorksheet.Cells[$sourceWorksheet.Dimension.Address].Value
}


# Fermeture des fichiers
Close-ExcelPackage -ExcelPackage $sourceWork -NoSave
Close-ExcelPackage -ExcelPackage $excelPackage