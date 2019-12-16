 param (
    [string]$username = "user",
    [string]$password = "pass",
    [string]$productversion = "pvers"
 )
powershell Set-ExecutionPolicy Unrestricted
$comAdmin = New-Object -comobject COMAdmin.COMAdminCatalog
$apps = $comAdmin.GetCollection("Applications")
$apps.Populate();

$newComPackageName = "1C83COMconnector"

$appExistCheckApp = $apps | Where-Object {$_.Name -eq $newComPackageName}

if($appExistCheckApp)
{
#$appExistCheckAppName = $appExistCheckApp.Value("Name")
"This COM+ Application already exists : $appExistCheckAppName"
$index1 = 0
foreach($appl in $apps) {
    if ($appl.Name -eq "1C83COMconnector") {
        $appColl.Remove($index1)
        $appColl.SaveChanges()
"deletion"
    }
    $index1++
}
}

#$comAdmin2 = New-Object -comobject COMAdmin.COMAdminCatalog
$apps = $comAdmin.GetCollection("Applications")
$apps.Populate();
$newApp1 = $apps.Add()
$newApp1.Value("Name") = $newComPackageName
$newApp1.Value("ApplicationAccessChecksEnabled") = 0 



<# Optional (to set to a specific Identify) #>
$newApp1.Value("Identity") = $username 
$newApp1.Value("Password") = $password 
$newApp1.Value("Description") = "com app for 1c"

$saveChangesResult = $apps.SaveChanges()
"Results of the SaveChanges operation : $saveChangesResult"

$comcntr = "C:\Program Files (x86)\1cv8\" + $productversion + "\bin\comcntr.dll"
<# $comAdmin = New-Object -comobject COMAdmin.COMAdminCatalog; #>
$comAdmin.InstallComponent("1C83COMconnector", $comcntr, $null, $null);

<# $components = $apps.GetCollection("Components",$newApp1.key)
$componentID = $component.Value("CLSID");
 
$RolesForComponent = $components.GetCollection("RolesForComponent",$component.Value("CLSID"))
$RolesForComponent.Populate(); #>

