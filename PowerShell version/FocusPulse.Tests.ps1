$ErrorActionPreference = 'Stop'

& (Join-Path $PSScriptRoot 'FocusPulse.ps1') -SelfTest
