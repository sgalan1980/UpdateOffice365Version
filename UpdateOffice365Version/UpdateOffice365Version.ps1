#
# Script.ps1
#
$Office365=0
$OfficeVersion = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*office 16*"} | fl version













ForEach($Office in $OfficeVersion)
	{
        "Producto $Office"
        if ($Office -Like "Office 16 Click-to-Run*")
           {
                 "Office 365 Instalado"
                 $Office365=1
                              
           }
      
        
        If ($Office -like "*Microsoft Office Standard 2010*")
            {
                "Se detecta $Office"
                 & cmd.exe /c "\\clublacosta\dfs\Packages\Office365\NewDeploy\Office2010Uninstall.Bat"
                  
             } 
             else
             {
                 if ($Office -Like "*Microsoft Office 2007*")
                 {
                        "Se detecta $Office"
                        #Invoke-Expression "E:\Scripts\Uninstall2013.bat"
                       & cmd.exe /c "\\clublacosta\dfs\Packages\Office365\NewDeploy\Office2007Uninstall.Bat"
                 }
				 else
                 {
                        if ($Office -Like "*Office*" -and $Office -like "*2003*")
                        {
                             "Se detecta $Office"
                               #Invoke-Expression "E:\Scripts\Uninstall2013.bat"
                             & cmd.exe /c "\\clublacosta\dfs\Packages\Office365\NewDeploy\Office2003Uninstall.Bat"
                        }
                 }
                    
                        
                    
             }

    }

if ($Office365 -eq 0)
{
        "Office 365 Sin instalar"
        Invoke-Expression "& `"C:\Program Files (x86)\Microsoft Office 365 ProPlus Installer\InstallOfficeProPlus.exe`""
}

