$Global:Current_Folder = split-path -parent $MyInvocation.MyCommand.Definition
$Date = get-date -format "dd-MM-yy_HHmm"
$CompName = $env:COMPUTERNAME
$Vendor = (gwmi win32_computersystemproduct).vendor

New-Alias XS Export-Services
New-Alias XD Export-Drivers
New-Alias XP Export-Process
New-Alias XB Export-BIOS
New-Alias XSo Export-Software
New-Alias XA Export-All


<#.Synopsis
	The Export-Services function allows you to export a services list from your computer. 
.DESCRIPTION
	Allow you to export a list of services from your computer.
	It will list each service with the following informations: Name, Caption, State, Start mode
	Services list can be export to the following format: CSV, XLSX, XML, HTML

.EXAMPLE
PS Root\> Export-Services -Path C:\ -csv
The command above will export a services list in CSV format in the folder C:\

.EXAMPLE
PS Root\> Export-Services -Path C:\ -xml
The command above will export a services list in XML format in the folder C:\

.EXAMPLE
PS Root\> Export-Services -Path C:\ -html
The command above will export a services list in HTML format in the folder C:\

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>	
	
Function Export-Services
{
[CmdletBinding()]
Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
        [string] $Path,
        [Switch] $CSV,
        [Switch] $XML,
        [Switch] $HTML		
      )
    
    Begin
    {		
		If ((-not $CSV) -and (-not $XML) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -csv to export in CSV or XLS format " -foregroundcolor "yellow"		 	
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}			
		ElseIf($CSV)
			{			
				Try
					{
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
						write-host ""		
						write-host "***************************************************"	
						write-host "Services will be exported in CSV and XLS format" -foregroundcolor "cyan"
					}

				Catch 
					{
						$Excel_value = $false 					
						write-host ""		
						write-host "***********************************************"	
						write-host "Services will be exported in CSV format" -foregroundcolor "cyan"
					}		
			}	
		ElseIf($XML)
			{
				write-host ""		
				write-host "**********************************************"	
				write-host "Services will be exported in XML format" -foregroundcolor "cyan"
			}	
		ElseIf($HTML)
			{
				$CSS_File = "$Current_Folder\Export-config.css" # CSS for HTML Export
				write-host ""		
				write-host "***************************************************"	
				write-host "Services will be exported in HTML format" -foregroundcolor "cyan"
			}				
    }
	

    Process
    {
		If($CSV)
			{
				$CSV_Services = "$Path\Export_Services.csv"				
				$win32_service = gwmi win32_service | select name, Caption, state, startmode 	
				$win32_service | Export-csv -path $CSV_Services -encoding UTF8 -notype -UseCulture 				
				If ($Excel_value -eq $True)
					{					
						$XLS_Services = "$Path\Export_Services.xlsx"				
						$xl = new-object -comobject excel.application
						$xl.visible = $False
						$xl.DisplayAlerts=$False
						$Workbook = $xl.workbooks.open($CSV_Services)
						$WorkSheet=$WorkBook.activesheet
						$table=$Workbook.ActiveSheet.ListObjects.add( 1,$Workbook.ActiveSheet.UsedRange,0,1)
						$WorkSheet.columns.autofit() | out-null
						$Workbook.SaveAs($XLS_Services,51)
						$Workbook.Saved = $True
						$xl.Quit()						
					}
			}		
		ElseIf($XML)
			{
				$XML_Services = "$Path\Export_Services.xml"				
				$win32_service = gwmi win32_service | select name, Caption, state, startmode 	
				$WIN32_service | export-clixml $XML_Services
			}	
		ElseIf($HTML)
			{
				$HTML_Services = "$Path\Export_Services.html"				
				$win32_service = gwmi win32_service | select name, Caption, state, startmode 	
				$services_list = "<p><span class=Main_Title>Services status list list on $CompName</span><br><span class=subtitle>This document has been updated on $date</span></p><br><br>"			
				$services_list_b = $win32_service | ConvertTo-HTML -Fragment
				$colorTagTable = @{Stopped = ' class="stopped">Stopped<';
								   Running = ' class="running">Running<'}
				$services_list = $services_list + $services_list_b				 
				$colorTagTable.Keys | foreach { $services_list = $services_list -replace ">$_<",($colorTagTable.$_) }
				ConvertTo-HTML  -body " $services_list" -CSSUri $CSS_File | 		
				Out-File -encoding ASCII $HTML_Services					
			}				
	}

    end
    {
		If($CSV)
			{			
				If ($Excel_value -eq $True)
					{	
						write-host "Services have been exported in CSV and XLS format" -foregroundcolor "green"
						write-host "***************************************************"							
					}
				Else
					{
						write-host "Excel seems to be not installed" -foregroundcolor "yellow"	
						write-host "Services have been exported in CSV format"						
					}
			}				
		ElseIf($XML)
			{
				write-host "Services have been exported in XML format" -foregroundcolor "green"
				write-host "***************************************************"	
			}	
			
		ElseIf($HTML)
			{
				write-host "Services have been exported in HTML format" -foregroundcolor "green"
				write-host "***************************************************"					
			}
	}			
				
	
}












<#.Synopsis
	The Export-Process function allow you to export a Process list from your computer. 
.DESCRIPTION
	Allow you to export a list of Process from your computer.
	It will list each service with the following informations: Name, Caption, State, Start mode
	Process list can be export to the following format: CSV, XLSX, XML, HTML

.EXAMPLE
PS Root\> Export-Process -Path C:\ -csv
The command above will export a Process list in CSV format in the folder C:\

.EXAMPLE
PS Root\> Export-Process -Path C:\ -xml
The command above will export a Process list in XML format in the folder C:\

.EXAMPLE
PS Root\> Export-Process -Path C:\ -html
The command above will export a Process list in HTML format in the folder C:\

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>

Function Export-Process
{
[CmdletBinding()]
Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
        [string] $Path,
        [Switch] $CSV,
        [Switch] $XML,
        [Switch] $HTML				
      )
    
    Begin
    {				
		If ((-not $CSV) -and (-not $XML) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"	
				write-host " !!! Use the switch -csv to export in CSV or XLS format " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}											
		ElseIf($CSV)
			{							
				Try
					{
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
						write-host ""		
						write-host "***************************************************"	
						write-host "Process will be exported in CSV and XLS format" -foregroundcolor "cyan"
					}

				Catch 
					{
						$Excel_value = $false 					
						write-host ""		
						write-host "***********************************************"	
						write-host "Process will be exported in CSV format" -foregroundcolor "cyan"
					}								
			}	
		ElseIf($XML)
			{
				write-host ""		
				write-host "**********************************************"	
				write-host "Process will be exported in XML format" -foregroundcolor "cyan"
			}	
		ElseIf($HTML)
			{
				$CSS_File = "$Current_Folder\Export-config.css" # CSS for HTML Export
				write-host ""		
				write-host "**********************************************"	
				write-host "Process will be exported in HTML format" -foregroundcolor "cyan"
			}				
    }

	
    Process
    {
		If($CSV)
			{
				$CSV_Process = "$Path\Export_Process.csv"		
				$win32_Process = gwmi win32_Process | select name, commandline 	
				$win32_Process | Export-csv -path $CSV_Process -encoding UTF8 -notype -UseCulture 	
				
				If ($Excel_value -eq $True)				
					{			
						$XLS_Process = "$Path\Export_Process.xlsx"
						$xl = new-object -comobject excel.application
						$xl.visible = $False
						$xl.DisplayAlerts=$False
						$Workbook = $xl.workbooks.open($CSV_Process)
						$WorkSheet=$WorkBook.activesheet
						$table=$Workbook.ActiveSheet.ListObjects.add( 1,$Workbook.ActiveSheet.UsedRange,0,1)
						$WorkSheet.columns.autofit() | out-null
						$Workbook.SaveAs($XLS_Process,51)
						$Workbook.Saved = $True
						$xl.Quit()							
					}						
			}		
		ElseIf($XML)
			{
				$XML_Process = "$Path\Export_Process.xml"				
				$win32_Process = gwmi win32_Process | select name, commandline 	
				$win32_Process | export-clixml $XML_Process
			}	
		ElseIf($HTML)
			{
				$HTML_Process = "$Path\Export_Process.html"				
				$win32_Process = gwmi win32_Process | select name, commandline 	
				$Process_list = "<p><span class=Main_Title>Process list list on $CompName</span><br><span class=subtitle>This document has been updated on $date</span></p><br><br>"						
				$Process_list_b = $win32_Process  | ConvertTo-HTML -Fragment
				$Process_list = $Process_list + $Process_list_b
				ConvertTo-HTML  -body " $Process_list" -CSSUri $CSS_File | 		
				Out-File -encoding ASCII $HTML_Process	
			}				
	}	
	

    end
    {
		If($CSV)
			{			
				If ($Excel_value -eq $True)
					{	
						write-host "Process have been exported in CSV and XLS format" -foregroundcolor "green"
						write-host "***************************************************"							
					}
				Else
					{
						write-host "Excel seems to be not installed" -foregroundcolor "yellow"	 
						write-host "Process have been exported in CSV format"						
					}
			}				
		ElseIf($XML)
			{
				write-host "Process have been exported in XML format" -foregroundcolor "green"
				write-host "***************************************************"	
			}	
			
		ElseIf($HTML)
			{
				write-host "Process have been exported in HTML format" -foregroundcolor "green"
				write-host "***************************************************"					
			}	
	}
	
}



















<#.Synopsis
	The Export-Hotfix function allow you to export a Hotfix list from your computer. 
.DESCRIPTION
	Allow you to export a list of Hotfix from your computer.
	It will list each Hotfix with the following informations: Hotfix ID, Description
	Hotfix list can be export to the following format: CSV, XLSX, XML, HTML

.EXAMPLE
PS Root\> Export-Hotfix -Path C:\ -csv
The command above will export a Hotfix list in CSV format in the folder C:\

.EXAMPLE
PS Root\> Export-Hotfix -Path C:\ -xml
The command above will export a Hotfix list in XML format in the folder C:\

.EXAMPLE
PS Root\> Export-Hotfix -Path C:\ -html
The command above will export a Hotfix list in HTML format in the folder C:\

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>

Function Export-Hotfix
{
[CmdletBinding()]
Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
        [string] $Path,
        [Switch] $CSV,
        [Switch] $XML,
        [Switch] $HTML				
      )
    
    Begin
    {		
		If ((-not $CSV) -and (-not $XML) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"	
				write-host " !!! Use the switch -csv to export in CSV or XLS format " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}											
		ElseIf($CSV)
			{				
				Try
					{
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
						write-host ""		
						write-host "***************************************************"	
						write-host "Hotfix will be exported in CSV and XLS format" -foregroundcolor "cyan"
					}

				Catch 
					{
						$Excel_value = $false 					
						write-host ""		
						write-host "***********************************************"	
						write-host "Hotfix will be exported in CSV format" -foregroundcolor "cyan"
					}							
			}	
		ElseIf($XML)
			{
				write-host ""		
				write-host "**********************************************"	
				write-host "Hotfix will be exported in XML format" -foregroundcolor "cyan"
			}	
		ElseIf($HTML)
			{
				$CSS_File = "$Current_Folder\Export-config.css" # CSS for HTML Export
				write-host ""		
				write-host "**********************************************"	
				write-host "Hotfix will be exported in HTML format" -foregroundcolor "cyan"
			}				
    }

	
    Process
    {
		If($CSV)
			{
				$CSV_Hotfix = "$Path\Export_Hotfix.csv"			
				$win32_quickfixengineering = gwmi win32_quickfixengineering | select hotfixid, description	
				$win32_quickfixengineering | Export-csv -path $CSV_Hotfix -encoding UTF8 -notype -UseCulture 
						
				If ($Excel_value -eq $True)				
					{					
						$XLS_Hotfix = "$Path\Export_Hotfix.xlsx"
						$xl = new-object -comobject excel.application
						$xl.visible = $False
						$xl.DisplayAlerts=$False
						$Workbook = $xl.workbooks.open($CSV_Hotfix)
						$WorkSheet=$WorkBook.activesheet
						$table=$Workbook.ActiveSheet.ListObjects.add( 1,$Workbook.ActiveSheet.UsedRange,0,1)
						$WorkSheet.columns.autofit() | out-null
						$Workbook.SaveAs($XLS_Hotfix,51)
						$Workbook.Saved = $True
						$xl.Quit()								
					}									
			}		
		ElseIf($XML)
			{
				$XML_Hotfix = "$Path\Export_Hotfix.xml"				
				$win32_quickfixengineering = gwmi win32_quickfixengineering | select hotfixid, description	
				$win32_quickfixengineering | export-clixml $XML_Hotfix				
			}	
		ElseIf($HTML)
			{				
				$HTML_Hotfix = "$Path\Export_Hotfix.html"	
				$win32_quickfixengineering = gwmi win32_quickfixengineering | select hotfixid, description					
				$Hotfix_list = "<p><span class=Main_Title>Hotfix list on $CompName</span><br><span class=subtitle>This document has been updated on $date</span></p><br><br>"				
				$Hotfix_list_b = $win32_quickfixengineering | ConvertTo-HTML -Fragment
				$Hotfix_list = $Hotfix_list + $Hotfix_list_b
				ConvertTo-HTML  -body " $Hotfix_list" -CSSUri $CSS_File | 		
				Out-File -encoding ASCII $HTML_Hotfix			
			}		
	}	

    end
    {
		If($CSV)
			{			
				If ($Excel_value -eq $True)
					{	
						write-host "Hotfix have been exported in CSV and XLS format" -foregroundcolor "green"
						write-host "***************************************************"							
					}
				Else
					{
						write-host "Excel seems to be not installed" -foregroundcolor "yellow"	
						write-host "Hotfix have been exported in CSV format"						
					}
			}				
		ElseIf($XML)
			{
				write-host "Hotfix have been exported in XML format" -foregroundcolor "green"
				write-host "***************************************************"	
			}	
			
		ElseIf($HTML)
			{
				write-host "Hotfix have been exported in HTML format" -foregroundcolor "green"
				write-host "***************************************************"					
			}		
	}
	
}

























<#.Synopsis
	The Export-Drivers function allows you to export a Drivers list from your computer. 
.DESCRIPTION
	Allow you to export a list of Drivers from your computer.
	It will list each service with the following informations: Device name, manufacturer, version, inf name
	Drivers list can be export to the following format: CSV, XLSX, XML, HTML

.EXAMPLE
PS Root\> Export-Drivers -Path C:\ -csv
The command above will export a Drivers list in CSV format in the folder C:\

.EXAMPLE
PS Root\> Export-Drivers -Path C:\ -xml
The command above will export a Drivers list in XML format in the folder C:\

.EXAMPLE
PS Root\> Export-Drivers -Path C:\ -html
The command above will export a Drivers list in HTML format in the folder C:\

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>

Function Export-Drivers
{
[CmdletBinding()]
Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
        [string] $Path,
        [Switch] $CSV,
        [Switch] $XML,
        [Switch] $HTML				
      )
    
    Begin
    {		
		If ((-not $CSV) -and (-not $XML) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"	
				write-host " !!! Use the switch -csv to export in CSV or XLS format " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}											
		ElseIf($CSV)
			{	
				Try
					{
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
						write-host ""		
						write-host "***************************************************"	
						write-host "Drivers will be exported in CSV and XLS format" -foregroundcolor "cyan"
					}

				Catch 
					{
						$Excel_value = $false 					
						write-host ""		
						write-host "***********************************************"	
						write-host "Drivers will be exported in CSV format" -foregroundcolor "cyan"
					}
			}	
		ElseIf($XML)
			{
				write-host ""		
				write-host "**********************************************"	
				write-host "Drivers will be exported in XML format" -foregroundcolor "cyan"
			}	
		ElseIf($HTML)
			{
				$CSS_File = "$Current_Folder\Export-config.css" # CSS for HTML Export
				write-host ""		
				write-host "**********************************************"	
				write-host "Drivers will be exported in HTML format" -foregroundcolor "cyan"
			}				
    }

	
    Process
    {
		If($CSV)
			{
				$CSV_Drivers = "$Path\Export_Drivers.csv"			
				$Win32_PnPSignedDriver = gwmi Win32_PnPSignedDriver | Select devicename, manufacturer, driverversion, infname | where-object {$_.infname -ne $null} 
				$Win32_PnPSignedDriver | Export-csv -path $CSV_Drivers -encoding UTF8 -notype -UseCulture 		

				If ($Excel_value -eq $True)								
					{			
						$XLS_Drivers = "$Path\Export_Drivers.xlsx"
						$xl = new-object -comobject excel.application
						$xl.visible = $False
						$xl.DisplayAlerts=$False
						$Workbook = $xl.workbooks.open($CSV_Drivers)
						$WorkSheet=$WorkBook.activesheet
						$table=$Workbook.ActiveSheet.ListObjects.add( 1,$Workbook.ActiveSheet.UsedRange,0,1)
						$WorkSheet.columns.autofit() | out-null
						$Workbook.SaveAs($XLS_Drivers,51)
						$Workbook.Saved = $True
						$xl.Quit()	
					}												
			}		
		ElseIf($XML)
			{
				$XML_Drivers = "$Path\Export_Drivers.xml"				
				$Win32_PnPSignedDriver = gwmi Win32_PnPSignedDriver | Select devicename, manufacturer, driverversion, infname | where-object {$_.infname -ne $null} 
				$Win32_PnPSignedDriver | export-clixml $XML_Drivers				
			}	
		ElseIf($HTML)
			{				
				$HTML_Drivers = "$Path\Export_Drivers.html"	
				$Win32_PnPSignedDriver = gwmi Win32_PnPSignedDriver | Select devicename, manufacturer, driverversion, infname | where-object {$_.infname -ne $null} 				
				$Drivers_list = "<p><span class=Main_Title>Drivers list on $CompName</span><br><span class=subtitle>This document has been updated on $date</span></p><br><br>"							
				$Drivers_list_b = $Win32_PnPSignedDriver |
						Select-Object devicename, manufacturer, driverversion, infname | 
						where-object {$_.infname -ne $null} | ConvertTo-HTML -Fragment
				$Drivers_list = $Drivers_list + $Drivers_list_b
				ConvertTo-HTML  -body " $Drivers_list" -CSSUri $CSS_File | 		
				Out-File -encoding ASCII $HTML_Drivers	
			}				
	}	
	

    end
    {
		If($CSV)
			{			
				If ($Excel_value -eq $True)
					{	
						write-host "Drivers have been exported in CSV and XLS format" -foregroundcolor "green"
						write-host "***************************************************"							
					}
				Else
					{
						write-host "Excel seems to be not installed" -foregroundcolor "yellow"	
						write-host "Drivers have been exported in CSV format"						
					}
			}				
		ElseIf($XML)
			{
				write-host "Drivers have been exported in XML format" -foregroundcolor "green"
				write-host "***************************************************"	
			}	
			
		ElseIf($HTML)
			{
				write-host "Drivers have been exported in HTML format" -foregroundcolor "green"
				write-host "***************************************************"					
			}					
	}
	
}




<#.Synopsis
	The Export-Software function allows you to export a Software list from your computer. 
.DESCRIPTION
	Allow you to export a list of Software from your computer.
	It will list each service with the following informations: Name, Version
	Software list can be export to the following format: CSV, XLSX, XML, HTML

.EXAMPLE
PS Root\> Export-Software -Path C:\ -csv
The command above will export a Software list in CSV format in the folder C:\

.EXAMPLE
PS Root\> Export-Software -Path C:\ -xml
The command above will export a Software list in XML format in the folder C:\

.EXAMPLE
PS Root\> Export-Software -Path C:\ -html
The command above will export a Software list in HTML format in the folder C:\

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>

Function Export-Software
{
[CmdletBinding()]
Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
        [string] $Path,
        [Switch] $CSV,
        [Switch] $XML,
        [Switch] $HTML				
      )
    
    Begin
    {		
		If ((-not $CSV) -and (-not $XML) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"	
				write-host " !!! Use the switch -csv to export in CSV or XLS format " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}											
		ElseIf($CSV)
			{		
				Try
					{
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
						write-host ""		
						write-host "***************************************************"	
						write-host "Software will be exported in CSV and XLS format" -foregroundcolor "cyan"
					}

				Catch 
					{
						$Excel_value = $false 					
						write-host ""		
						write-host "***********************************************"	
						write-host "Software will be exported in CSV format" -foregroundcolor "cyan"
					}					
			}	
		ElseIf($XML)
			{
				write-host ""		
				write-host "**********************************************"	
				write-host "Software will be exported in XML format" -foregroundcolor "cyan"
			}	
		ElseIf($HTML)
			{
				$CSS_File = "$Current_Folder\Export-config.css" # CSS for HTML Export
				write-host ""		
				write-host "**********************************************"	
				write-host "Software will be exported in HTML format" -foregroundcolor "cyan"
			}				
    }

	
    Process
    {
		If($CSV)
			{
				$CSV_Software = "$Path\Export_Software.csv"							
				$win32_product = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion 
				$win32_product | Export-csv -path $CSV_Software -encoding UTF8 -notype -UseCulture 				

				If ($Excel_value -eq $True)												
					{			
						$XLS_Software = "$Path\Export_Software.xlsx"
						$xl = new-object -comobject excel.application
						$xl.visible = $False
						$xl.DisplayAlerts=$False
						$Workbook = $xl.workbooks.open($CSV_Software)
						$WorkSheet=$WorkBook.activesheet
						$table=$Workbook.ActiveSheet.ListObjects.add( 1,$Workbook.ActiveSheet.UsedRange,0,1)
						$WorkSheet.columns.autofit() | out-null
						$Workbook.SaveAs($XLS_Software,51)
						$Workbook.Saved = $True
						$xl.Quit()
					}															
			}		
		ElseIf($XML)
			{
				$XML_Software = "$Path\Export_Software.xml"				
				$win32_product = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion 
				$win32_product | export-clixml $XML_Software	
			}	
		ElseIf($HTML)
			{				
				$HTML_Software = "$Path\Export_Software.html"	
				$win32_product = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion 				
				$Software_list = "<p><span class=Main_Title>Software list on $CompName</span><br><span class=subtitle>This document has been updated on $date</span></p><br><br>"					
				$Software_list_b = $win32_product  | ConvertTo-HTML -Fragment
				$Software_list = $Software_list + $Software_list_b
				ConvertTo-HTML  -body " $Software_list" -CSSUri $CSS_File | 		
				Out-File -encoding ASCII $HTML_Software				
			}				
	}	

    end
    {
		If($CSV)
			{			
				If ($Excel_value -eq $True)
					{	
						write-host "Software have been exported in CSV and XLS format" -foregroundcolor "green"
						write-host "***************************************************"							
					}
				Else
					{
						write-host "Excel seems to be not installed" -foregroundcolor "yellow"	
						write-host "Software have been exported in CSV format"						
					}
			}				
		ElseIf($XML)
			{
				write-host "Software have been exported in XML format" -foregroundcolor "green"
				write-host "***************************************************"	
			}	
			
		ElseIf($HTML)
			{
				write-host "Software have been exported in HTML format" -foregroundcolor "green"
				write-host "***************************************************"					
			}		
	}
	
}












<#.Synopsis
	The Export-BIOS function allows you to export a BIOS settings list from your computer. 
.DESCRIPTION
	Allow you to export a list of BIOS settings from your computer.
	BIOS settings list can be export to the following format: CSV, XLSX, XML, HTML

.EXAMPLE
PS Root\> Export-Software -Path C:\ -csv
The command above will export a BIOS settings list in CSV format in the folder C:\

.EXAMPLE
PS Root\> Export-Software -Path C:\ -xml
The command above will export a BIOS settings list in XML format in the folder C:\

.EXAMPLE
PS Root\> Export-Software -Path C:\ -html
The command above will export a BIOS settings list in HTML format in the folder C:\

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>

Function Export-BIOS
{
[CmdletBinding()]
Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
        [string] $Path,
        [Switch] $CSV,
        [Switch] $XML,
        [Switch] $HTML				
      )
    
    Begin
    {		
		If ((-not $CSV) -and (-not $XML) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"	
				write-host " !!! Use the switch -csv to export in CSV or XLS format " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}											
		ElseIf($CSV)
			{	
				Try
					{
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
						write-host ""		
						write-host "***************************************************"	
						write-host "BIOS settings will be exported in CSV and XLS format" -foregroundcolor "cyan"
					}

				Catch 
					{
						$Excel_value = $false 					
						write-host ""		
						write-host "***********************************************"	
						write-host "BIOS settings will be exported in CSV format" -foregroundcolor "cyan"
					}		
			}	
		ElseIf($XML)
			{
				write-host ""		
				write-host "*******************************************************"	
				write-host "BIOS settings will be exported in XML format" -foregroundcolor "cyan"
			}	
		ElseIf($HTML)
			{
				$CSS_File = "$Current_Folder\Export-config.css" # CSS for HTML Export
				write-host ""		
				write-host "*******************************************************"	
				write-host "BIOS settings will be exported in HTML format" -foregroundcolor "cyan"
			}				
    }
	
    Process
    {		
		If ($Vendor -eq "LENOVO")
			{	
				If($CSV)
					{
						$CSV_BIOS_Settings = "$Path\BIOS_Settings.csv"						
						$BIOS_WMI = gwmi -class Lenovo_BiosSetting -namespace root\wmi  | select currentsetting | Where {$_.CurrentSetting -ne ""} |
						select @{label = "Name"; expression = {$_.currentsetting.split(",")[0]}} , 
						@{label = "active_value"; expression = {$_.currentsetting.split(",*;[")[1]}}	
						$BIOS_WMI | Export-csv -path $CSV_BIOS_Settings -encoding UTF8 -notype -UseCulture  
						
						If ($Excel_value -eq $True)												
							{			
								$XLS_BIOS_Settings = "$Path\BIOS_Settings.xlsx"					
								$xl = new-object -comobject excel.application
								$xl.visible = $False
								$xl.DisplayAlerts=$False
								$Workbook = $xl.workbooks.open($CSV_BIOS_Settings)
								$WorkSheet=$WorkBook.activesheet
								$table=$Workbook.ActiveSheet.ListObjects.add( 1,$Workbook.ActiveSheet.UsedRange,0,1)
								$WorkSheet.columns.autofit() | out-null
								$Workbook.SaveAs($XLS_BIOS_Settings,51)
								$Workbook.Saved = $True
								$xl.Quit()	
							}																						
					}		
				ElseIf($XML)
					{
						$XML_BIOS_Settings = "$Path\BIOS_Settings.xml"
						
						$BIOS_WMI = gwmi -class Lenovo_BiosSetting -namespace root\wmi  | select currentsetting | Where {$_.CurrentSetting -ne ""} |
						select @{label = "Name"; expression = {$_.currentsetting.split(",")[0]}} , 
						@{label = "active_value"; expression = {$_.currentsetting.split(",*;[")[1]}}							
						
						$BIOS_Col = gwmi win32_bios 
						ForEach ($objItem in $BIOS_Col) 
							{
								$BIOS_Ver = $objItem.SMBIOSBIOSVersion 
							}		
						$BIOS_WMI | export-clixml $XML_BIOS_Settings						
					}	
				ElseIf($HTML)
					{				
						$HTML_BIOS_Settings = "$Path\BIOS_Settings.html"
						
						$BIOS_WMI = gwmi -class Lenovo_BiosSetting -namespace root\wmi  | select currentsetting | Where {$_.CurrentSetting -ne ""} |
						select @{label = "Name"; expression = {$_.currentsetting.split(",")[0]}} , 
						@{label = "active_value"; expression = {$_.currentsetting.split(",*;[")[1]}}							
												
						$BIOS_Col = gwmi win32_bios 
						foreach ($objItem in $BIOS_Col) 
							{
								$BIOS_Ver = $objItem.SMBIOSBIOSVersion 
							}

						$Title = "<p><span class=Main_Title>BIOS Settings list on $CompName with BIOS Version $BIOS_Ver</span><br><span class=subtitle>This document has been updated on $date</span></p><br>"						
						$BIOS_WMI | ConvertTo-HTML  -body " $Title<br>$BIOS_WMI" -CSSUri $CSS_File | 		
						Out-File -encoding ASCII $HTML_BIOS_Settings	
					}				
		}	
	}
    end
    {
		If($CSV)
			{			
				If ($Excel_value -eq $True)
					{	
						write-host "BIOS settings have been exported in CSV and XLS format" -foregroundcolor "green"
						write-host "***************************************************"							
					}
				Else
					{
						write-host "Excel seems to be not installed" -foregroundcolor "yellow"	
						write-host "BIOS settings have been exported in CSV format"						
					}
			}				
		ElseIf($XML)
			{
				write-host "BIOS settings have been exported in XML format" -foregroundcolor "green"
				write-host "***************************************************"	
			}	
			
		ElseIf($HTML)
			{
				write-host "BIOS settings have been exported in HTML format" -foregroundcolor "green"
				write-host "***************************************************"					
			}		
	}
	
}












<#.Synopsis
	The Export-All function allows you to export a list with the following configurations: Services, Process, Drivers, Software, BIOS settings. 
.DESCRIPTION
	Allow you to export a list of services, software, drivers, process and BIOS settings from your computer.
	Lists can be export to the following format: CSV, XLSX, XML, HTML

.EXAMPLE
PS Root\> Export-All -Path C:\ -csv
The command above will export a list of services, software, drivers, process and BIOS settings in CSV format in the folder C:\

.EXAMPLE
PS Root\> Export-All -Path C:\ -xml
The command above will export a list of services, software, drivers, process and BIOS settings in XML format in the folder C:\

.EXAMPLE
PS Root\> Export-All -Path C:\ -html
The command above will export a list of services, software, drivers, process and BIOS settings in HTML format in the folder C:\

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>

Function Export-All
	{
		[CmdletBinding()]
		Param(
				[Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
				[string] $Path,
				[Switch] $CSV,
				[Switch] $XML,
				[Switch] $HTML		
			  )	
	
		Begin
			{		
				If ((-not $CSV) -and (-not $XML) -and (-not $HTML))
					{
						write-host ""		
						write-host "*******************************************************************"	
						write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"
						write-host " !!! Use the switch -csv to export in CSV or XLS format " -foregroundcolor "yellow"		
						write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
						write-host "*******************************************************************"					
					}
				Else
					{			
						write-host ""		
						write-host "********************************************************************************************************************"	
						write-host "Services, hotfix, drivers, software and BIOS settings settings will be exported" -foregroundcolor "Cyan"
						write-host "********************************************************************************************************************"	
					}
			}	
		Process
			{				
				If($CSV)
					{						
						Export-Services -path $Path -CSV
						Export-Drivers -path $Path -CSV
						Export-Software -path $Path -CSV
						Export-Hotfix -path $Path -CSV						
						Export-Process -path $Path -CSV
						Export-BIOS	-path $Path	-CSV 		
					}
				If($XML)
					{						
						Export-Services -path $Path -XML
						Export-Drivers -path $Path -XML
						Export-Software -path $Path -XML
						Export-Hotfix -path $Path -XML												
						Export-Process -path $Path -XML
						Export-BIOS	-path $Path	-XML 
					}
				If($HTML)
					{						
						Export-Services -path $Path -html
						Export-Drivers -path $Path -html
						Export-Software -path $Path -html
						Export-Hotfix -path $Path -html												
						Export-Process -path $Path -html
						Export-BIOS	-path $Path	-html 	
					}	
			}		
	
	}
