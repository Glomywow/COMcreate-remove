$comCatalog = New-Object -ComObject COMAdmin.COMAdminCatalog
$appColl = $comCatalog.GetCollection("Applications")
$appColl.Populate()
 
$app = $appColl | where {$_.Name -eq "1C83COMconnector"}
$index1 = 0
foreach($appl in $appColl) {
    if ($appl.Name -eq "1C83COMconnector") {
        $appColl.Remove($index1)
        $appColl.SaveChanges()
"deleting"
    }
    $index1++
}
#$app.delete()
#$compColl = $appColl.GetCollection("Components", $app.Key)
#$compColl.Populate()
 
#$index = 0
#foreach($component in $compColl) {
#    if ($component.Name -eq "1C83COMconnector") {
#        $compColl.Remove($index)
#        $compColl.SaveChanges()
#    }
#    $index++
#}
