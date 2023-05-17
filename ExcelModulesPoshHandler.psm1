# Adding namespaces:
using namespace Microsoft.VBE.Interop
using namespace Microsoft.Office.Interop.Excel

<#

  .SYNOPSIS
  A short one-line action-based description, e.g. 'Tests if a function is valid'

  .DESCRIPTION 
  Exports the Excel VBAProject members from the passed Excel file/s into the folders.

  .NOTES
  Information or caveats about the function e.g. 'This function is not supported in Linux'

  .LINK
  Specify a URI to a help page, this will show when Get-Help -Online is used.

  .EXAMPLE
  Test-MyTestFunction -Verbose
  Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
  .INPUTS

  System.String[]
  
  .OUTPUTS
  System.IO.File

#> 
function Export-ExcelModules {
  [CmdletBinding()]
  Param (

    # Specifies the path to one or more Excel files. Wildcard characters are permitted.
    [Parameter(Mandatory, Position=0)]
    [string[]]$Path,

    # Specifies the path to a resource. Unlike Path , the value of the LiteralPath parameter is used exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters, enclose it in single quotation marks. Single quotation marks tell PowerShell not to interpret any characters as escape sequences.
    [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)][Alias("PSPath", "LP")]
    [string[]]$LiteralPath,

    # Specifies the path to where the modules will be exported. The default is the current directory.
    [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [string]$Destination = [System.Environment]::CurrentDirectory,

    # Specifies the name of the Export folder. The default is the Name of the Excel file + "_" + the Name of the VBProject.
    [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)][Alias("Name", "EFN")]
    [string]$ExportFolderName

  )

  begin {

    # Add the Excel Application Namespace:
    Add-type -AssemblyName Microsoft.VBE.Interop;
    Add-type -AssemblyName Microsoft.Office.Interop.Excel;

    # Clear the Error catch:
    $Error.Clear();

    # Convert the paths to PSPaths:
    $Path = Convert-Path -Path $Path;
    $Destination = Convert-Path -Path $Destination;

    # Create Excel object:
    $excel = new-object -TypeName "Microsoft.Office.Interop.Excel.ApplicationClass";

  } process {

    try {
      
      foreach ($PathMember in $Path) {

        # Get the file name from the full path:
        $FileName = Split-Path -Path $PathMember -Leaf;
        
        # Check if the $Path contains an Excel file:
        if ( $FileName -notlike "*.xl*" ) { Write-Host " '$FileName' is not an Excel file! " -ForegroundColor "Red" ; Continue  };
        
        # Open the Excel file:
        $eFile = $excel.Workbooks.Open("$PathMember");

        #region Checks, if the oppened Excel file contains a VBA project:

        if ($eFile.HasVBProject) {
          ## If yes:

          Write-Verbose -Message "Excel file '$FileName' has VBA project.";

          $vbProject = $eFile.VBProject;

          # Set default ExportFolderName:
          if ( $ExportFolderName -eq "" ) { $ExportFolderName = "$FileName`_$($vbProject.Name)" };
          
          # Generate a Code hierarchy folders:
          $RootFolder =  New-Item -Path "$Destination\$ExportFolderName" -ItemType "Directory" -Force;
          $UserFormsPath = New-Item -Path $RootFolder -Name "UserForms" -ItemType Directory -Force;
          $ModulesPath = New-Item -Path $RootFolder -Name "Modules" -ItemType Directory -Force
          $ClassModulesPath = New-Item -Path $RootFolder -Name "ClassModules" -ItemType Directory -Force;
          $OthersPath = New-Item -Path $RootFolder -Name "OtherObjects" -ItemType Directory -Force;

          # Final path variable:
          $ExportPath = "";
          
          foreach ($component in $vbProject.VBComponents) {

            # Sort the current Component by its type:
            switch ([Microsoft.Vbe.Interop.vbext_ComponentType]$component.Type) {
              
              ( [Microsoft.Vbe.Interop.vbext_ComponentType]::vbext_ct_MSForm ) { 
                Write-Verbose -Message "'$($component.name)' component type is: A UserForm.";
                
                $ExportPath = Convert-Path -Path $UserFormsPath;
                $ExportPath += "\$($component.name).frm";
                
              };
              
              ( [Microsoft.Vbe.Interop.vbext_ComponentType]::vbext_ct_StdModule ) {
                Write-Verbose -Message "'$($component.name)' component type is: A Module";
                
                $ExportPath = Convert-Path -Path $ModulesPath;
                $ExportPath += "\$($component.name).bas";

              };

              ( [Microsoft.Vbe.Interop.vbext_ComponentType]::vbext_ct_ClassModule ) {              
                Write-Verbose -Message "'$($component.name)' component type is: A ClassModule.";

                $ExportPath = Convert-Path -Path $ClassModulesPath;
                $ExportPath += "\$($component.name).cls";

              };

              Default {
                Write-Verbose -Message "'$($component.name)' component type is: An other type of an object.";
                
                $ExportPath = Convert-Path -Path $OthersPath;
                $ExportPath += "\$($component.name).cls";

              }

            }
            
            # Export the current Component:
            $component.Export("$ExportPath");

          } # end of the foreach loop.
        
        } else { Write-Verbose -Message "Excel file '$FileName' has no VBAProject." };
        
        #endregion

        # Close the current Worksheet:
        $eFile.Close();

      }

    } catch {
      Write-Error -Message "$_";
    }

  } end {

    # Quit the Excel app:
    $excel.Quit()

    return;
  }
};

#region Configuration handling ( CRUD / NGSR ):

function New-ExcelModulesConfiguration {
  [CmdletBinding()]
  param ()
  
}

function Get-ExcelModulesConfiguration {
  param(

    # Get all existing configurations.
    [Parameter()]
    [switch]$All

  )
  return Read-ConfigFile;
};

function Set-ExcelModulesConfiguration {
  [CmdletBinding()]
  param (

    # Name/s of the parameter/s.
    [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
    [string[]]$Name,

    # New value/s of the parameter/s.
    [Parameter(Mandatory, ValueFromPipelineByPropertyName)][AllowNull()][AllowEmptyString()]
    [string[]]$Value,

    # Prompts you for confirmation before running the function.
    [Parameter(ValueFromPipelineByPropertyName)][Alias('cf')]
    [switch]$Confirm

  )
  
  begin{

    # 
    # if ($Name.Count -ne $Value.Count) {
    #   Throw
    # }
    [PSCustomObject]$currentConfigBody = Read-ConfigFile;

  } process {
    $parameterIndex = 0;
    foreach ($configName in $Name) {

      <# $configName is the current item #>
      $currentConfigBody.PSObject.Properties.foreach({
        if ( $_.Name -eq $configName ) {

          Write-Host -Object "$configName`: We have a match.";
          Write-Host "BEFORE: $($_.Name) = $($_.value)";
          Write-Host "AFTER: $configName = $($value)";

          # If $true, user will be prompted to confirm his modification:
          if ($Confirm) { 


          } else { $currentConfigBody.$configName = $Value; }
        } else { Write-Warning -Message "Parameter '$configName' is not in the current configuration."; }
        $parameterIndex++;
      })

    }

  } end { return $currentConfigBody }
};

function Remove-ExcelModulesConfiguration {
  param()
  
}


#endregion Configuration handling.

#region Implementation:

<#
  .DESCRIPTION
  Reads User-defined or Default configuration.
#>
function Read-ConfigFile {
  param(
    # Read Default Configuration file instead of User-defined.
    [Parameter()]
    [switch]$Default
  )

  # Handle the switch parameter:
  if ( $Default ) { $configFile = "DefaultConfig.json" } else { $configFile = "UserConfig.json" };

  # Return the configuration:
  return ( Get-Content -Path ".\Files\Configuration\$configFile" | ConvertFrom-Json );
  
}

#endregion implementation.

$ExportFuntions = @(
  "Export-ExcelModules",
  "New-ExcelModulesConfiguration",
  "Get-ExcelModulesConfiguration",
  "Set-ExcelModulesConfiguration",
  "Remove-ExcelModulesConfiguration"
);

Export-ModuleMember -Function $ExportFuntions;