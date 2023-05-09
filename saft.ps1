# Coding languague: English | EN
# User displayed language: Portuguese | PT
# Script created: 27/03/2023
# Script updated: 09/04/2023

Add-Type -AssemblyName System.Windows.Forms

$var = $True #this variable serves the purpose to stay or exit the while loop
$filePath = "C:\SAFT" #Folder is pre defined

# Asks the User for the needed information
Function User-Info()
{
    $year = Read-Host -Prompt "Introduza o ano"
    $month = Read-Host -Prompt "Introduza o mês"
    $securePasswd = Read-Host -Prompt "Introduza a password" -assecurestring
    return $year, $month, $securePasswd
}

# Opens the file explorer showing .xml or .jar files to select and grabs the path and name
function Get-FileName($filePath, $extension)
{
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $filePath
    $OpenFileDialog.filter = "Documento (*." + $extension + ") | *." + $extension
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.FileName
}

Clear Host
Write-Host ":::::: Envio do ficheiro SAFT ::::::`n"
$year, $month, $securePasswd = User-Info

while ($var -eq $true)
{
    # JAR file choosing ||||||||||||
    Write-Host "`nSelecione o ficheiro JAR"
    $jarFile = Get-FileName $filePath "jar"

    # Checks if a 'cancel' or 'close' was selected instead by checking is the variable is empty
    if(!$jarFile)
    {
        Write-Host "`nScript terminado"
        exit 0
    }else
    {
        #checks if the file selected is a .jar type of file
        $fileExtension = [System.IO.Path]::GetExtension($jarFile)
        if($fileExtension -ne ".jar")
        {
            Write-Host "Ficheiro selecionado não é um .jar! Escolha outro ficheiro!"
            continue
        }
    }
    # XML file choosing ||||||||||||
    Write-Host "`nSelecione o ficheiro XML"
    $xmlFile = Get-FileName $filePath "xml"

    # Checks if a 'cancel' or 'close' was selected instead by checking is the variable is empty
    if(!$xmlFile)
    {
        Write-Host "`nScript terminado"
        exit 0
    }else
    {
        #checks if the file selected is a .xml type of file
        $fileExtension = [System.IO.Path]::GetExtension($xmlFile)
        if($fileExtension -ne ".xml")
        {
            Write-Host "Ficheiro selecionado não é um .xml! Escolha outro ficheiro!"
            continue
        }else
        {
            $var = $False
        }
    }
}

Write-Host "`nFicheiro JAR selecionado: $jarFile"
Write-Host "Ficheiro XML selecionado: $xmlFile"

#Gets the current date to use as the file name
$logDate = (Get-Date).tostring("yyyy-MM-dd_HH-mm")

# Transform secure password to plain text to be able to show the variable $script result
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePasswd)
$unsecurePassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

$extension = ".xml"
$script = "& `"C:\Program Files\Java\jre1.8.0_341\bin\java`" -jar `"$jarFile`" -n 507236335/4 -p $unsecurePassword -a $year -m $month -op enviar -i `"$xmlFile`"  -o `"$filePath\Resultado_$logDate$extension`""

# Cleans the password from being 'visible'
[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

Invoke-Expression $script
Write-Host "`nScript executado: $script"
