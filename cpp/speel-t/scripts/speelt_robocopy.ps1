$exitCode = (Start-Process -FilePath 'robocopy' -ArgumentList @PSBoundParameters -PassThru -Wait).ExitCode

if ($exitCode -lt 7) {
    exit 0
}

exit $exitCode