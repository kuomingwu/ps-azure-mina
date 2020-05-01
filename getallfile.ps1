$pathPrefix = "C:\Users\ki59920\Desktop\Mina\";
Connect-AzureRmAccount

$list = @() ;

function GetResourceBySubscriptionId(){
    param(
        [String] $SubscriptionId
    )
    
    Select-AzureRmSubscription -SubscriptionId $SubscriptionId
    $CurrentResourceList = Get-AzureRmResource
    $resourceList = @();
    $CurrentResourceList | foreach {
        $obj = @{ 
             'SubscriptionId' = $_.SubscriptionId 
             'ResourceName' = $_.ResourceName 
             'ResourceType' = $_.ResourceType
             'ResourceId' = $_.ResourceId 
        }
        
        $resourceList += $obj ;
        
    }
    
    return $resourceList ;


}

function exportToExcel ($list){
    $excel = New-Object -ComObject excel.application
    $excel.visible = $False
    $workbook = $excel.Workbooks.Add()
    $diskSpacewksht= $workbook.Worksheets.Item(1)
    $diskSpacewksht.Name = "All resource"
    $diskSpacewksht.Cells.Item(2,8) = 'Mina Design'
    $diskSpacewksht.Cells.Item(2,8).Font.Size = 18
    $diskSpacewksht.Cells.Item(2,8).Font.Bold=$True
    $diskSpacewksht.Cells.Item(2,8).Font.Name = "Cambria"
    $diskSpacewksht.Cells.Item(2,8).Font.ThemeFont = 1
    $diskSpacewksht.Cells.Item(2,8).Font.ThemeColor = 4
    $diskSpacewksht.Cells.Item(2,8).Font.ColorIndex = 55
    $diskSpacewksht.Cells.Item(2,8).Font.Color = 8210719


    $diskSpacewksht.Cells.Item(3,1) = 'SubscriptionId'
    $diskSpacewksht.Cells.Item(3,2) = 'ResourceType'
    $diskSpacewksht.Cells.Item(3,3) = 'ResourceId'
    $diskSpacewksht.Cells.Item(3,4) = 'ResourceName'

    $index = 4 ;
    for($i = 0 ; $i -lt $list.Count ; $i++){
       $from = $index + $i ;    
       $diskSpacewksht.Cells.Item($from,1) = $list[$i].SubscriptionId
       $diskSpacewksht.Cells.Item($from,2) = $list[$i].ResourceType
       $diskSpacewksht.Cells.Item($from,3) = $list[$i].ResourceId
       $diskSpacewksht.Cells.Item($from,4) = $list[$i].ResourceName

    }

    
    #excel.DisplayAlerts = 'False'
    $ext=".xlsx"
    $name = (Get-Date -UFormat %s) -Replace("[,\.]\d*", "") ;
    $path="$pathPrefix$name$ext"
    $workbook.SaveAs($path) 
    $workbook.Close
    $excel.DisplayAlerts = 'False'
    $excel.Quit()


}




$info = Get-AzureRmSubscription | Select -ExpandProperty "SubscriptionId"

$info | foreach { 
    
    $resourceList = GetResourceBySubscriptionId($_) 
    $list += $resourceList ;
}



exportToExcel($list);






