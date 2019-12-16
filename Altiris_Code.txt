function Get-MetaData
{
PARAM
(
        [PARAMETER(Mandatory=$true)]
        [string]$path = ""
    ,   [PARAMETER( Mandatory=$true,
                    HelpMessage="extension IE. .mp3, .txt or .mov")]
        [string]$type = ""#'mp3'
    ,   [switch]$recurse
)
    if ($recurse)
    {
        $LPath = Get-ChildItem -Path $path -Directory -Recurse
    } else {$LPath = $path}

    $DirectoryCount = 1
    $RetrievedMetadata = $true
    $OutputList = New-Object 'System.Collections.generic.List[psobject]'

    Foreach ($pa in $LPath)
    {
        $shell = New-Object -ComObject shell.application

        if ($recurse)
        {
            $objshell = $shell.NameSpace($pa.FullName)
        } else {$objshell = $shell.NameSpace($pa)}
        #Build data list
        $count = 0

        #Filter on filetype
        $filter = $objshell.items() | where {$_.path -match $type}
        foreach ($file in $filter)
        {
            if ($RetrievedMetadata)
            {
                # Build metanumbers
                Write-Verbose "Building MetaIndex for filetype $type"
                $Metanumbers = New-Object -TypeName 'System.Collections.Generic.List[int]'
                for ($a = 0 ; $a -le 400;$a++)
                {
                    if ($objshell.GetDetailsOf($file,$a))
                    {
                        $Metanumbers.Add([int]$a)
                    }    
                    if ($objshell.GetDetailsOf($file,$a) -ne "")
                    {
                        $Metanumbers.Add([int]$a)
                    }
                }
                $RetrievedMetadata = $false
                Write-Verbose "$($Metanumbers.Count) entries in MetaIndex for filetype $type"    
            }

            $count++
            $CurrentDirectory = Get-ChildItem -Path $file.path
            try{
            Write-Progress -Activity " Getting metadata from $DirectoryCount/$($LPath.count) $($CurrentDirectory[0].DirectoryName)" -Status "Working on $count/$($filter.count) - $($file.Name)" -PercentComplete (($count / $filter.count) * 100) -ErrorAction stop
            }catch{}

            #Build Hashtable for each file
            $Hash = @{}
            foreach ($nr in $Metanumbers)
            {
                $hash += @{$($objshell.getDetailsOf($objshell.items, $nr))  = 
                   $($objshell.getDetailsOf($File, $nr))} 
            }
            
            $Hash.Remove("")
            $FileMetaData = New-Object -TypeName PSobject -Property $hash
            $OutputList.Add($FileMetaData)
            $hash = $null
        }
        $DirectoryCount++        
    }
    Write-Verbose "MetaData for $($OutputList.count) files found"

    $OutputList | Export-Clixml -Path "C:\Users\trevor.k.jones2.ctr\Desktop\Data.xml"

    return $OutputList | Format-List
}

Function Get-Folder($initialDirectory)

{
    # Set-ExecutionPolicy Bypass
    # Set-ExecutionPolicy Unrestricted
    # Set-ExecutionPolicy RemoteSigned

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select a folder"
    $foldername.rootfolder = "MyComputer"

    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

# IF YOU WANT TO MANUALLY SELECT A FILE, UNCOMMENT THE BELOW METHOD AND COMMENT OUT LINE 122
 $filePath = Get-Folder

# GETS CURRENT LOCATION SCRIPT IS BEING RAN IN I.E THE PUBLISH FOLDER FOR WEB APPLICATIONS
# $filePath = $PSScriptRoot

Get-MetaData -path $filePath -type dll
