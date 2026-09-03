# Marks the built MSI as UAC compliant (Summary Information Word Count bit 3, value 8).
# Without this flag Windows Installer requests elevation even for per-user installs.
# Run as a post-build event of the Setup project: mark-uac-compliant.ps1 "<path to .msi>"
param([Parameter(Mandatory)][string]$MsiPath)

$installer = New-Object -ComObject WindowsInstaller.Installer
$db = $installer.OpenDatabase($MsiPath, 1)   # 1 = msiOpenDatabaseModeTransact
$summary = $db.SummaryInformation(1)
$wordCount = [int]$summary.Property(15)
if (($wordCount -band 8) -eq 0) {
    $summary.Property(15) = $wordCount -bor 8
    $summary.Persist()
    $db.Commit()
    Write-Host "mark-uac-compliant: Word Count $wordCount -> $($wordCount -bor 8) in '$MsiPath'"
} else {
    Write-Host "mark-uac-compliant: already set ($wordCount) in '$MsiPath'"
}
