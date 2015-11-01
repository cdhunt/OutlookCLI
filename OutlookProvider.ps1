Add-type -assembly "Microsoft.Office.Interop.Outlook" | Out-Null 
$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]  
$outlook = new-object -comobject outlook.application 
$namespace = $outlook.GetNameSpace("MAPI") 
$inbox = $namespace.getDefaultFolder($olFolders::olFolderInBox)

root {
    script Inbox {    
        $inbox.Items
    }

    $inbox.Folders | foreach-object {
        $subFolder = $_
        $content = Split-Path $subFolder.folderpath -Leaf
        
        script "$content" {
            $subFolder.Items
        }.GetNewClosure();
    }

}
