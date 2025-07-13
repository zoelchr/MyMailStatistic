# buildexe.ps1
Import-Module PS2EXE          # falls nicht schon automatisch geladen

$ps2exeParams = @{
    InputFile  = '.\MailStatistic.ps1'
    OutputFile = '.\MailStatistic.exe'
    NoConsole  = $false
    #IconFile   = '.\chart.ico'
    Version    = '1.0.0'
    Title      = 'Mail Statistic'
    Parameters = 'YearsBack 10'
}

Invoke-PS2EXE @ps2exeParams
Write-Host 'Fertig – EXE erzeugt.'