#-----------------------------------------------------------------------------------------------------------------------------------
# Fonction : Code qui permet de recuperer les fichiers xls ou xlsx de repartitions ou se trouvent les droits d'auteur de l'artiste  
#            utilise un fichier de configuration Json ou se trouve tous les parametres d'initialisation
# Entree   : username et password d'un compte SACEM
# Sortie   : les fichiers xls ou xlsx de repartitions
# Auteur   : Theo Breton
# date     : 5/13/2021
# Version  : V12
#-----------------------------------------------------------------------------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#------------------------------------------------------------------
# lis le json
#----------------------------------------------------------------

$json = Get-Content -Raw -Path "$PSScriptRoot\config-for-sacemv2.json" | ConvertFrom-Json

 #---------------------------------------------------------------
 # create variable login
 #---------------------------------------------------------------

$login = Invoke-WebRequest -uri $json.'link-createurs-sacem'
cls

#log in

write-host "start"

$login.Forms[0].Fields.Keys

write-host "end"
Write-host "Form id     : " $login.Forms[0].Id
Write-host "Form Action : " $login.Forms[0].Action
Write-host "Form Method : " $login.Forms[0].Method


#username and password  to log in are temporarily used brut.

#-----------------------
# positionne les fields
#-----------------------

$login.Forms[0].Fields.username = $json.username
$login.Forms[0].Fields.password = $json.Password

Foreach ( $b in $login.Forms[0].Fields.Keys )
        {
          write-host "Fields      : " $b " / " $login.Forms[0].Fields[$b]
        }

write-host ""


$mainPage = Invoke-WebRequest -uri $json.'link-for-login' -Body $login -Method POST -SessionVariable var1

#-----------------------------
# initialisation after log in
#-----------------------------

#-----------------------------
# nom de l'artiste
#-----------------------------


$artiste= $mainPage.parsedhtml.body.GetElementsByClassName("txtUpp mbm")[0].innertext

write-host "artiste = "$artiste

#-------------------------------
# premier et dernier repartition
#-------------------------------

$mainPage = Invoke-WebRequest -uri $json.'link-des-tableau' -Method GET  -WebSession $var1
  
#Write-host "Result      : " $? "," $mainpage.StatusCode "," $mainpage.StatusDescription
#Write-host "" 

$dropdown= $mainPage.scripts.Item(2).outerText 
$CharArray =$dropdown.Split("'")

$i1max=0

$max = 0

$min = 999

$tab1=@()

$coad

$count=0

$i= 0

foreach($item in $CharArray)

{ #write-host "id=" $item
  # check Coadid for each item

    if ( $item.contains('coadId'))
 
            {
                $str3 = $item

                $str3=$str3.Replace(",","")
                $str3=$str3.Replace("activeFilters:","")
                $str3=$str3.Replace("{","")
                $str3=$str3.Replace("coadId:","")
                $str3=$str3.Replace("  ","")

                write-host "id="$str3
         
            }

      $i = $i+1  

# check Repartition for each item

    if ($item.Contains("R\u00E9partition"))

    {     

        if ($item.Contains("{1}") -eq $false)

            {
                    $item

                    $previousitem

                    $previousitem2

                    $i1max=$i1max+1
           
                    if ($previousitem2 -gt $max)               

                        {

                            $max = $previousitem2

                         }

 

                    if ($previousitem2 -le $min)

                        {

                        $min = $previousitem2

                        }         
            }

# End if Repartition
    }  
 
    $previousitem2=$previousitem

    $previousitem=$item 

# end Foreach

  } 

#-----------------------------------   
# affichage de tableau de min et max
#-----------------------------------

$icoadid= $str3
write-host "coaID ="$icoadid

write-host "min = "$min
write-host "max = "$max

#-----------------------------------------------------------------------------------------
# DOWNLOAD DE REPARTITION
#-----------------------------------------------------------------------------------------

#-------------------------------------------------------------------------------------------------------------------------------------
# download repartition 631-647
# repartition en brut for now
# example: https://nouveau-feuillet.sacem.fr/download/excel-repartition?dataType=0&category=cdeFamilleTypeUtil&coad=478385&period=631
#-------------------------------------------------------------------------------------------------------------------------------------    

#----------------------------------
# telechargement des repartitions
#----------------------------------

write-host "minimum=" $min
write-host "maximum=" $max
$icoaid= $str3 
$imin = [int]$min
$imax = [int]$max


for ($r=$imin; $r -le $imax  ; $r++)
{
    if($r -lt 647)
     {
     
      
     $url2 = $json.'linkforoldrepartition' + "dataType=" + $json.'datatype' + '&' + "category="+$json.'category' + "&" + 'coad='+$icoaid + "&" + 'period=' + $r
    
        write-host $url2

        $output4 = "$PSScriptRoot\__02-fichierxls64\$r.xls"
    
        Invoke-WebRequest -Uri $url2 -OutFile $output4 -WebSession $var1
    
       
         Write-host "Result      : " $? "," $login.StatusCode ,"," $r 

         $b64 = Get-Content "$PSScriptRoot\__02-fichierxls64\$r.xls"

         $filename = "$PSScriptRoot\$r.xls"

         $bytes = [Convert]::FromBase64String($b64)

         [IO.File]::WriteAllBytes($filename, $bytes)

         Move-Item -Path "$PSScriptRoot\$r.xls" -Destination "$PSScriptRoot\__01-toutes-les-repartitions" -Force
     }
    
#--------------------------------------------------------
# Download repartition 647-652 envoie les requetes posts
#--------------------------------------------------------

    If ($r -ge 647)
   
    { 

        for ($r = $imin; $r -le $imax;$r++)    
        {
              
            write-host $r  
               
            write-host "status=" $mainpage.StatusCode
            
            $output= "$PSScriptRoot\__03-zip\$r.zip" 

            Invoke-WebRequest -Uri $url -OutFile $output -WebSession $var1

            Expand-Archive -path "$PSScriptRoot\__03-zip\$r.zip" -DestinationPath "$PSScriptRoot\__01-toutes-les-repartitions" -Force

            $url = $json.'link-for-new-repartition-download' + 'OM-' + $r +'-'+ $icoaid + '.zip'

            write-host $url

            write-host "urlzip=" $url
            

        }   
    }
 }


#--------------------------------
# Merge all excel sheets
#--------------------------------
 
$objexcel  = New-Object -ComObject excel.application
$objexcel2 = New-Object -ComObject excel.application

$objexcel.visible=$false
$objexcel2.visible=$false

$ExcelFiles = Get-ChildItem -Path "$PSScriptRoot\__01-toutes-les-repartitions"

# -----------------------------------
# List all XLXS Files
# -----------------------------------

foreach($ExcelFile in $ExcelFiles)
{
 write-host "File Name  : " $ExcelFile.FullName
}

# -----------------------------------
# Create XLXS  Global File
# -----------------------------------
$workbookglobal=$objexcel2.Workbooks.add()
$worksheetglobal=$workbookglobal.Sheets.Item("sheet1")

$countline = 0 

foreach($ExcelFile in $ExcelFiles)
        {
        
        $workbook=$objexcel.Workbooks.Open($ExcelFile.FullName)
        $worksheet = $workbook.sheets.item(1)

        Write-Host "Sheet Name :" $worksheet.Name -ForegroundColor Cyan

        $str1 =""

        $maxrow     = ($worksheet.UsedRange.Rows).count

        $maxCol     = ($worksheet.UsedRange.Columns).count

        write-host "Max Row    : " ($worksheet.UsedRange.Rows).count -ForegroundColor Green

        write-host "Max Col    : " ($worksheet.UsedRange.Columns).count -ForegroundColor Green

        # only first line on first file , after begin at 2
        if ( $countline -eq 0 )
            { $istart = 1 }
        else
            { $istart = 2  } 

        for ($j = $istart ; $j -le $maxrow ; $j++ )

            {
                    for ($i =1 ; $i -le $maxcol ; $i++ )
                        {
                            $str1 = $str1 + $worksheet.cells.Item($j, $i).text + ","  
                            $currentline = $countline + $j 
                            $worksheetglobal.cells.Item( $currentline , $i).value= $worksheet.cells.Item($j, $i).text
                              
                        }

                     $str1 = $str1 + "`r`n"
            }

            $countline =  $countline + $maxrow-1

            write-host $str1

            $objexcel.Workbooks.Close()
 
        }

         #$Everysheet.Copy($Worksheet)
         #$Everyexcel.Close()

$Workbookglobal.SaveAs("$PSScriptRoot\__05-global\$artiste")

$objexcel.Workbooks.close()
$objexcel.Quit()

$objexcel2.Workbooks.close()
$objexcel2.Quit()

#------------------------------------
# Removes all useless repartitions
#------------------------------------

Remove-Item -path "$PSScriptRoot\__01-toutes-les-repartitions\*.xls"
Remove-Item -path "$PSScriptRoot\__01-toutes-les-repartitions\*.xlsx"

#------------------------------------------
# This line removes the useless zip files
# Removes all useless .xls files
#------------------------------------------
   
Remove-Item -path "$PSScriptRoot\__03-zip\*.zip" 
Remove-Item -path "$PSScriptRoot\__02-fichierxls64\*.xls"
