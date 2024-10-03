@{
    # Assemblies that must be loaded prior to importing this module
    RequiredAssemblies = $null

    # Script module or binary module file associated with this manifest.
    RootModule         = 'OTS-PSModules.psm1'

    # Version number of this module.
    ModuleVersion      = '0.2'

    # ID used to uniquely identify this module
    GUID               = 'f531c8ef-a6fd-4448-a1f8-45252cbb894c'

    # Author of this module
    Author             = 'Ehren Holgersson'

    # Company or vendor of this module
    CompanyName        = 'Olympus Technical Support'

    # Copyright statement for this module
    Copyright          = 'c 2024 All rights reserved.'

    # Description of the functionality provided by this module
    Description        = 'Various PS Functions created for management and troubleshooting'

    # Functions to export from this module
    FunctionsToExport  = @(
        'OTS-OfficeRepair',
        'Get-LatestOfficeVersion'
    )

    # Aliases to export from this module
    AliasesToExport    = @()

    # Cmdlets to export from this module
    CmdletsToExport    = @()

    FileList           = @(
        '.\OTS-PSModules.psd1',
        '.\OTS-PSModules.psm1',
        '.\OTS-PSModules\Public\OTS-OfficeRepair.ps1',
        '.\OTS-PSModules\Private\Get-LatestOfficeVersion.ps1',
        '.\OTS-PSModules\Private\Remove-FromString.ps1'
    )

    # Private data to pass to the module specified in RootModule/ModuleToProcess
    
    # Minimum version of the Windows PowerShell engine required by this module
    # PowerShellVersion = ''

    # Name of the Windows PowerShell host required by this module
    # PowerShellHostName = ''

    # Minimum version of the Windows PowerShell host required by this module
    # PowerShellHostVersion = ''

    # Minimum version of Microsoft .NET Framework required by this module
    # DotNetFrameworkVersion = ''

    # Minimum version of the common language runtime (CLR) required by this module
    # CLRVersion = ''

    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''

    # Modules that must be imported into the global environment prior to importing this module
    # RequiredModules = @()

    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    # ScriptsToProcess = @()

    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()

    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()

    # List of all modules packaged with this module
    # ModuleList = @()

    # List of all files packaged with this module
    # Variables to export from this module
    #VariablesToExport = '*'

    # HelpInfo URI of this module
    # HelpInfoURI = ''

    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''

}