# Shared application functions. Dot-sourced by RRFQComparer.ps1.
# Keep this file free of application startup code.

function Get-AppConfig {
    $configPath = Join-Path $script:AppRoot 'config\suppliers.json'
    if (Test-Path -LiteralPath $configPath) {
        return Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }

    return [pscustomobject]@{
        Suppliers = @(
            [pscustomobject]@{ Name = 'Ben'; Aliases = @('ben', 'бен') },
            [pscustomobject]@{ Name = 'BMZ'; Aliases = @('bmz', 'бмз') },
            [pscustomobject]@{ Name = 'HNN'; Aliases = @('hnn') },
            [pscustomobject]@{ Name = 'CPR'; Aliases = @('cpr') },
            [pscustomobject]@{ Name = 'Компэл'; Aliases = @('compel', 'компэл', 'компел') },
            [pscustomobject]@{ Name = 'Промэлектроника'; Aliases = @('promelec', 'promelektronika', 'промэлектроника') },
            [pscustomobject]@{ Name = 'Др.'; Aliases = @('other', 'др', 'другой') }
        )
        OtherSupplier = 'Др.'
        VatDivisor = 1.22
    }
}

function Get-SupplierNames {
    @($script:Config.Suppliers | ForEach-Object { $_.Name })
}

function Get-SupplierFromFileName {
    param([string]$Path)

    $name = ([IO.Path]::GetFileNameWithoutExtension($Path)).ToLowerInvariant()
    foreach ($supplier in $script:Config.Suppliers) {
        foreach ($alias in @($supplier.Aliases)) {
            if ($name.Contains(([string]$alias).ToLowerInvariant())) {
                return [string]$supplier.Name
            }
        }
    }

    return [string]$script:Config.OtherSupplier
}
