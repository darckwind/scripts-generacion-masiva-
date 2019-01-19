#get arrray of pendrives
$array = (Get-WmiObject Win32_Volume -Filter "DriveType='2'").DriveLetter

#system for multi copy items
foreach($b in $array)
    {
        #template css
        Copy-Item -Path C:\Users\$env:UserName\Desktop\Parche-problema-PRP\css\template.css -Destination $b"\Servidor_Portable\xampp\htdocs\lsinicia\evaluacion\upload\templates\INICIA_beta\css\" -Force
        Copy-Item -Path C:\Users\$env:UserName\Desktop\Parche-problema-PRP\css\template.css -Destination $b"\Servidor_Portable\xampp\htdocs\lsinicia\evaluacion\tmp\assets\7edb3ec4\css\" -Force
        Copy-Item -Path C:\Users\$env:UserName\Desktop\Parche-problema-PRP\css\template.css -Destination $b"\Servidor_Portable\xampp\htdocs\lsinicia\evaluacion\tmp\assets\20939c18\css\" -Force
        Copy-Item -Path C:\Users\$env:UserName\Desktop\Parche-problema-PRP\css\template.css -Destination $b"\Servidor_Portable\xampp\htdocs\lsinicia\evaluacion\tmp\assets\67432d9a\css\" -Force
        Copy-Item -Path C:\Users\$env:UserName\Desktop\Parche-problema-PRP\css\template.css -Destination $b"\Servidor_Portable\xampp\htdocs\lsinicia\evaluacion\tmp\assets\ab50146a\css\" -Force

        #template js
        Copy-Item -Path C:\Users\$env:UserName\Desktop\Parche-problema-PRP\js\template.js -Destination $b"\Servidor_Portable\xampp\htdocs\lsinicia\evaluacion\upload\templates\INICIA_beta\scripts\" -Force
        Copy-Item -Path C:\Users\$env:UserName\Desktop\Parche-problema-PRP\js\template.js -Destination $b"\Servidor_Portable\xampp\htdocs\lsinicia\evaluacion\tmp\assets\ab50146a\scripts\" -Force
        Copy-Item -Path C:\Users\$env:UserName\Desktop\Parche-problema-PRP\js\template.js -Destination $b"\Servidor_Portable\xampp\htdocs\lsinicia\evaluacion\tmp\assets\7edb3ec4\scripts\" -Force
        Copy-Item -Path C:\Users\$env:UserName\Desktop\Parche-problema-PRP\js\template.js -Destination $b"\Servidor_Portable\xampp\htdocs\lsinicia\evaluacion\tmp\assets\20939c18\scripts\" -Force
        Copy-Item -Path C:\Users\$env:UserName\Desktop\Parche-problema-PRP\js\template.js -Destination $b"\Servidor_Portable\xampp\htdocs\lsinicia\evaluacion\tmp\assets\67432d9a\scripts\" -Force

    }

cls

echo "parche terminado UwU"

echo "inicio reporte"


#create folder for reports
$ruta= "C:\Users\$env:UserName\Desktop\Reportes-parche"

if(!(Test-Path -Path $ruta )){
    New-Item -ItemType directory -Path $ruta
}


cls

#get disk data
#$disks=Get-WmiObject -Class Win32_LogicalDisk 


$xl=New-Object -ComObject "Excel.Application" 
 
$wb=$xl.Workbooks.Add()
$ws=$wb.ActiveSheet
 
$cells=$ws.Cells
 
$cells.item(1,1)="Disk Drive Report"
$cells.item(1,1).font.bold=$True
$cells.item(1,1).font.size=18
 
#define some variables to control navigation
$row=3
$col=1
 
#insert column headings
"Drive","SizeGB","FreespaceGB","UsedGB","%Free","%Used","directory","files","Status" | foreach {
    $cells.item($row,$col)=$_
    $cells.item($row,$col).font.bold=$True
    $col++
}
 
foreach ($arrays in $array) {
    $drive = (Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='$arrays'")
    $row++
    $col=1
    $cells.item($Row,$col)=$drive.DeviceID
    $col++
    $cells.item($Row,$col)=$drive.Size
    $cells.item($Row,$col).NumberFormat="0"
    $col++
    $cells.item($Row,$col)=$drive.Freespace
    $cells.item($Row,$col).NumberFormat="0.00"
    $col++
    $cells.item($Row,$col)=($drive.Size - $drive.Freespace)
    $cells.item($Row,$col).NumberFormat="0"
    $col++
    $cells.item($Row,$col)=($drive.Freespace/$drive.size)
    $cells.item($Row,$col).NumberFormat="0.00%"
    $col++
    $cells.item($Row,$col)=($drive.Size - $drive.Freespace) / $drive.size
    $cells.item($Row,$col).NumberFormat="0.00%"
    $col++
    $cells.item($Row,$col)=(Get-ChildItem -Recurse -Force -Directory $arrays).Count
    $col++
    $cells.item($Row,$col)=(Get-ChildItem -Recurse -Force -File $arrays).Count

    #archivos
    if(($cells.item($Row,8).Value() -eq 58891) -or ($cells.item($Row,8).Value() -eq 59198)){
        #directorios
        if(($cells.item($Row,7).Value() -eq 6243) -or ($cells.item($Row,7).Value() -eq 6292)){
            #espacio usado
            if(($cells.item($Row,4).Value() -eq 1445249024) -or ($cells.item($Row,4).Value() -eq 1456893952)){
                $col++
                $cells.item($Row,$col)="pass"
            }else{
                $col++
                $cells.item($Row,$col)="fail"
            }
        }else{
            $col++
            $cells.item($Row,$col)="fail"
        }
    }else{
    $col++
    $cells.item($Row,$col)="fail"
    }

   
}
 

$xl.Visible=$True
 
$filepath="C:\Users\$env:UserName\Desktop\Reportes-parche\Reporte"

$count = (Get-ChildItem -Recurse -Force -File "C:\Users\$env:UserName\Desktop\Reportes-parche").Count
sleep 2
 
if ($filepath) {
    $wb.SaveAs($filepath+"-"+$count) 
}

$row=3
$col=1

$wb.Save()
$wb.Close()

$xl.Quit()