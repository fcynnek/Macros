#Using PowerShell to delete folders in SharePoint Online 

Get-PSSnapin -Registered | Add-PSSnapin -Passthru
Add-PSSnapIn -Name Microsoft.Exchange, Microsoft.Windows.AD

$web = Get-SPWeb -Identity "https://unitednations.sharepoint.com/sites/DOS-HRM-SAS"    

$listname = $web.GetList("https://unitednations.sharepoint.com/sites/DOS-HRM-SAS/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FDOS%2DHRM%2DSAS%2FShared%20Documents%2FGS%20and%20Related%20Tests")    

function DeleteFiles   
{    
  param($folderUrl)    
  $folder = $web.GetFolder($folderUrl)      
  foreach ($file in $folder.Files)   
  {    
    # Delete file by deleting parent SPListItem    
    Write-Host("DELETED FILE: " + $file.name)    
    $list.Items.DeleteItemById($file.Item.Id)    
  }    

}    

# Delete root files    

DeleteFiles($listname.RootFolder.Url)    

# Delete files in folders    

foreach ($folder in $listname.Folders)   
{    
  DeleteFiles($folder.Url)    
}    

# Delete folders    

foreach ($folder in $list.Folders)   
{    
  try   
  {  
    Write-Host("DELETED FOLDER: " + $folder.name)    
    $list.Folders.DeleteItemById($folder.ID)    
  }    
  catch   
  {  
    Write-Host(“Deletion of parent folder already deleted this folder”)    
  }    
} 

