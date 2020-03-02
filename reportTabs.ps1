###############################################################################################
#
#
#					        Functions
#
#
###############################################################################################


Function hostAlarms($cluster){
    $esx_all = $cluster | Get-VMHost | Get-View
    $Report=@()
    foreach ($esx in $esx_all){
        foreach($triggered in $esx.TriggeredAlarmState){
            If ($triggered.OverallStatus -like "red" ){
                $lineitem={} | Select Name, AlarmInfo
                $alarmDef = Get-View -Id $triggered.Alarm
                $lineitem.Name = $esx.Name
                $lineitem.AlarmInfo = $alarmDef.Info.Name
                $Report+=$lineitem
            } 
        }
    }
    $Report |Sort Name | export-csv "c:\temp\ESX-Host-Red-Alarms.csv" -notypeinformation -useculture
    Invoke-item "c:\temp\ESX-Host-Red-Alarms.csv"
}

##############################################################################
#                  Add Snapins Required to Run the Scripts
############################################################################## 
## Add snapins

Add-PSSnapin VMware.VimAutomation.Core -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Xml.Linq -ErrorAction SilentlyContinue

##############################################################################
#                   Set the Size of the Window when Running
############################################################################## 
$pshost = Get-Host
$pswindow = $pshost.ui.rawui

$newsize = $pswindow.buffersize
$newsize.height = 3000
$newsize.width = 150
$pswindow.buffersize = $newsize

$newsize = $pswindow.windowsize
$newsize.height = 50
$newsize.width = 150
$pswindow.windowsize = $newsize

##############################################################################
#                   Set the Global Variables
############################################################################## 

$vcenterReport = "vCenter Report"
$WriteOutList = $Null
$myFileDate = Get-Date -format yyyyMMdd_HH_mm
$ScriptPath = (Split-Path ((Get-Variable MyInvocation).Value).MyCommand.Path)
$outputpath = Join-Path -Path $ScriptPath -ChildPath "Output"

$ReportFile = "$outputpath" + "\$vcenterReport $myfiledate.html"
New-Item -ItemType file $ReportFile -Force
$ErrorLog = "$outputpath" + "\$vcenterReport error log $myfiledate.log"
New-Item -ItemType file $ErrorLog -Force

Disconnect-VIServer -Confirm:$false -ErrorAction:SilentlyContinue

Do {
    $myCred = Get-Credential
    $Response3  = Read-Host 'Press Y to continue or any key to re-enter the password'
} until ($Response3 -eq 'y')

$vCenters = #"vcenter1.mydomain.com",
            #"vcenter2.mydomain.com"

$vcID = 0
$stArray  = "["
$vcObjArr = @()

forEach ($vCenterFQDN in $vCenters){
            $vcID++
            $vcObject = new-object System.Object
            $vcObject | add-Member -type NoteProperty -name VCFQDN -value $vCenterFQDN
            $vcObject | add-Member -type NoteProperty -name PAGEID -value $vcID
            $vcObjArr += $vcObject
            $stArray += "'page-$vcID',"
}

$stArray = "$stArray"
$stArray = $stArray.Substring(0,$stArray.length-1)
$stArray += "]"
$stArray = "$stArray"

###############################################################################################
#
#
#					                    Create the HTML for the Report
#
#
###############################################################################################

Add-Content $ReportFile "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Strict//EN' https://protect-eu.mimecast.com/s/4-afCE87jh6vrAZCNkcwo?domain=w3.org'>"
Add-Content $ReportFile "<html xmlns='https://protect-eu.mimecast.com/s/uMbQCGv7lsLDxyzT7pCKO?domain=w3.org'>"
Add-Content $ReportFile "<head>"
Add-Content $ReportFile "<style type='text/css'>"
Add-Content $ReportFile "div.content {"
Add-Content $ReportFile "    border: #48f solid 3px;"
Add-Content $ReportFile "    clear: left;"
Add-Content $ReportFile "    padding: 1em;"
Add-Content $ReportFile "    font-family: Tahoma;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "div.content.inactive {"
Add-Content $ReportFile "   display: none;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc {"
Add-Content $ReportFile "    height: 2em;"
Add-Content $ReportFile "    list-style: none;"
Add-Content $ReportFile "    margin: 0;"
Add-Content $ReportFile "    padding: 0;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc a {"
Add-Content $ReportFile "    background: #bdf url(tabs.gif);"
Add-Content $ReportFile "    color: #008;"
Add-Content $ReportFile "    display: block;"
Add-Content $ReportFile "    float: left;"
Add-Content $ReportFile "    height: 2em;"
Add-Content $ReportFile "    padding-left: 10px;"
Add-Content $ReportFile "    text-decoration: none;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc a:hover {"
Add-Content $ReportFile "    background-color: #3af;"
Add-Content $ReportFile "    background-position: 0 -120px;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc a:hover span {"
Add-Content $ReportFile "    background-position: 100% -120px;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc li {"
Add-Content $ReportFile "    float: left;"
Add-Content $ReportFile "    margin: 0 1px 0 0;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc li a.active {"
Add-Content $ReportFile "    background-color: #48f;"
Add-Content $ReportFile "    background-position: 0 -60px;"
Add-Content $ReportFile "    color: #fff;"
Add-Content $ReportFile "    font-weight: bold;"
Add-Content $ReportFile "    font-family: Tahoma;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc li a.active span {"
Add-Content $ReportFile "    background-position: 100% -60px;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "ol#toc span {"
Add-Content $ReportFile "    background: url(tabs.gif) 100% 0;"
Add-Content $ReportFile "    display: block;"
Add-Content $ReportFile "    line-height: 2em;"
Add-Content $ReportFile "    padding-right: 10px;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "</style>"
Add-Content $ReportFile "<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />"
Add-Content $ReportFile "<title>$reportDescripton</title>"
Add-Content $ReportFile "</head>"
Add-Content $ReportFile "<body>"
Add-Content $ReportFile "<body style=font-family:Tahoma>"
Add-Content $ReportFile "<h1>$reportDescripton</h1>"
Add-Content $ReportFile "<ol id='toc'>"  

###############################################################################################
#			                     Create a Tab Per vCenter 
###############################################################################################

forEach ($singObj in $vcObjArr){
    $pageID = $singObj.PAGEID
    $vcFQDN = $singObj.VCFQDN
    Add-Content $ReportFile "<li><a href='#page-$pageID'><span>$vcFQDN Info</span></a></li>" 
}

###############################################################################################
#			                     END - Create a Tab Per vCenter 
###############################################################################################

Add-Content $ReportFile "</ol>"
Add-Content $ReportFile "<style>BODY{background-color:Linen;}TABLE{font-family: Tahoma;border-width: 2px;border-style: solid;
border-color: black;border-collapse: collapse;}TH{border-width: 2px;padding: 2px;border-style: solid;border-color: black;
background-color:DodgerBlue}TD{border-width: 2px;padding: 2px;border-style: solid;border-color: black;background-color:PowderBlue}</style>"

###############################################################################################
#			                       Create a Page Per vCenter
###############################################################################################

forEach ($singObj in $vcObjArr){
    $pageID = $singObj.PAGEID
    $vcFQDN = $singObj.VCFQDN
    
    Add-Content $ReportFile "<div class='content' id='page-$pageID'>"
 
    disconnect-viserver -Confirm:$false -ErrorAction:SilentlyContinue
    connect-viserver -server $vcFQDN -credential $myCred 

    $startDTM = (Get-Date)
    ##enter content

###############################################################################################
#			                     Get vCenter Information
###############################################################################################

    $vcsi = Get-View -Server $vcFQDN ServiceInstance
    $vcenterVersion = $vcsi.content.About.Version
    $build  = $vcsi.content.About.Build

    $clusters = get-cluster
    $numberClusters = $clusters.count

    $vms = get-vm
    $numberVMs = $vms.count

    $vmHosts = get-vmhost
    $numberHosts = $vmHosts.count

    New-Object -TypeName PSObject -Property @{
        Version         = $vcenterVersion
        Build           = $build
        NumberClusters  = $numberClusters
        NumberVMs       = $numberVMs
        NumberHosts     = $numberHosts
    } | Select Version,Build,NumberClusters,NumberVMs,NumberHosts | Sort-Object -Property Version | ConvertTo-Html -Fragment | Out-File $OutputPath\$vcFQDN'info'.html

    Add-Content $ReportFile "<H2>$vcFQDN General Information</H2>"
    Get-Content $OutputPath\$vcFQDN'info'.html | Add-Content $ReportFile
    Remove-Item $OutputPath\$vcFQDN'info'.html 

###############################################################################################
#			                     Get Cluster Information
###############################################################################################

    foreach($cluster in $clusters){

        $clusterView = $cluster | get-view
        $clusterHA = $clusterView.Configuration.DasConfig.Enabled
            if ($clusterHA -like "True"){
                $clusterHA = "HA Enabled"
            } else {
                $clusterHA = "HA Disabled"
            }
        $admissionControl = $clusterView.Configuration.DasConfig.AdmissionControlEnabled
            if ($admissionControl -like "True"){
                $admissionControl = "Admission Control Enabled"
            } else {
                $admissionControl = "Admission Control Disabled"
        }


        $failureResponse = $clusterView.Configuration.DasConfig.DefaultVmSettings.RestartPriority
        $restartPriority = $clusterView.Configuration.DasConfig.DefaultVmSettings.RestartPriority
            if ($failureResponse -like "disabled"){
                $failureResponse = "Disabled"
            } else {
                $failureResponse = "Restart VMs"
        }

        $failoverCapacity = $clusterView.Configuration.DasConfig.AdmissionControlPolicy.GetTYpe().Name
        
        switch($failoverCapacity){
                'ClusterFailoverHostAdmissionControlPolicy' {$failoverCapacity = "Dedicated Failover Hosts"}
                'ClusterFailoverResourcesAdmissionControlPolicy' {$failoverCapacity = "Cluster Resource Percentage"}
                'ClusterFailoverLevelAdmissionControlPolicy' {$failoverCapacity = "Slot Policy"}
        }

        New-Object -TypeName psobject -Property @{
            HAEnabled                       = $clusterHA
            DefineHostFailoverCapacityBy    = $failoverCapacity
            AdmissionControlEnabled         = $admissionControl
            HostFailureResponse             = $failureResponse
            VMRestartPriority               = $restartPriority

        } | Select HAEnabled, DefineHostFailoverCapacityBy, AdmissionControlEnabled, HostFailureResponse, VMRestartPriority | Sort-Object -Property HAEnabled | ConvertTo-Html -Fragment | Out-File $OutputPath\$vcFQDN'HAinfo'.html


        Add-Content $ReportFile "<H2>$vcFQDN HA Information</H2>"
        Get-Content $OutputPath\$vcFQDN'HAinfo'.html | Add-Content $ReportFile
        Remove-Item $OutputPath\$vcFQDN'HAinfo'.html 

        $esx_all = $cluster | Get-VMHost | Get-View
        $esx_all | %{
            New-Object -TypeName PSObject -Property @{
                ESXHostName                     =  $_.Name
                ConnectionState                  = $_.Runtime.ConnectionState
                PowerState                       = $_.Runtime.PowerState
                MaintenanceMode                  = $_.Runtime.inMaintenanceMode
                BootTime                         = $_.Runtime.BootTime
        } | Select ESXHostName,ConnectionState,PowerState,MaintenanceMode,BootTime
    } | Sort-Object -Property ESXHostName | ConvertTo-Html -Fragment | Out-File $OutputPath\$cluster'HostConnectivity'.html
        
       
        Add-Content $ReportFile "<H2>$cluster Host Connectivity and Uptime</H2>"
        Get-Content $OutputPath\$cluster'HostConnectivity'.html | Add-Content $ReportFile
        Add-Content $ReportFile "<br><br>"                
        Remove-Item $OutputPath\$cluster'HostConnectivity'.html        
    }
  
###############################################################################################
#			                     Get Host Information
###############################################################################################

###############################################################################################
#			                     Get Network Information
###############################################################################################

###############################################################################################
#			                     Get Datastore Information
###############################################################################################

$allDatastore = get-datastore | Get-View


foreach($ds in $allDatastore){
      
}
    
    
    #Add-Content $ReportFile "$vmhost Basic Information"
    #Add-Content $ReportFile "<br><br>"
    #Get-Content $OutputPath"\$vmhost basicInfo.html" | Add-Content $ReportFile
    #Remove-Item $OutputPath"\$vmhost basicInfo.html"
    #Add-Content $ReportFile "<br><br>"
    #Add-Content $ReportFile "$vmhost Network Information"
    #Add-Content $ReportFile "<br><br>"
    #Get-Content $hostNicFile | Add-Content $ReportFile
    #Remove-Item $hostNicFile
    #Add-Content $ReportFile "<br><br>"
    #Add-Content $ReportFile "$vmhost Key Information"
    #Add-Content $ReportFile "<br><br>"
    #Get-Content $OutputPath\$vmhost'advancedInfo'.html | Add-Content $ReportFile
    #Remove-Item $OutputPath\$vmhost'advancedInfo'.html

    Add-Content $ReportFile "</div>"

}

###############################################################################################
#
#
#					         Finish the Report and write the Javascript
#
#
###############################################################################################

Add-Content $ReportFile "</body>"
Add-Content $ErrorLog "Formatting HTML report with JAVA"
Add-Content $ReportFile "<script type='text/javascript'>"
Add-Content $ReportFile "// Wrapped in a function so as to not pollute the global scope."
Add-Content $ReportFile "var activatables = (function () {"
Add-Content $ReportFile "// The CSS classes to use for active/inactive elements."
Add-Content $ReportFile "var activeClass = 'active';"
Add-Content $ReportFile "var inactiveClass = 'inactive';"
Add-Content $ReportFile "  "
Add-Content $ReportFile "var anchors = {}, activates = {};"
Add-Content $ReportFile "var regex = /#([A-Za-z][A-Za-z0-9:._-]*)$/;"
Add-Content $ReportFile "  "
Add-Content $ReportFile "// Find all anchors (<a href='#something'>.)"
Add-Content $ReportFile "var temp = document.getElementsByTagName('a');"
Add-Content $ReportFile "for (var i = 0; i < temp.length; i++) {"
Add-Content $ReportFile "     var a = temp[i];"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   // Make sure the anchor isn't linking to another page."
Add-Content $ReportFile "   if ((a.pathname != location.pathname &&"
Add-Content $ReportFile "       '/' + a.pathname != location.pathname) ||"
Add-Content $ReportFile "       a.search != location.search) continue;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   // Make sure the anchor has a hash part."
Add-Content $ReportFile "   var match = regex.exec(a.href);"
Add-Content $ReportFile "   if (!match) continue;"
Add-Content $ReportFile "   var id = match[1];"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   // Add the anchor to a lookup table."
Add-Content $ReportFile "   if (id in anchors)"
Add-Content $ReportFile "       anchors[id].push(a);"
Add-Content $ReportFile "   else"
Add-Content $ReportFile "       anchors[id] = [a];"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "// Adds/removes the active/inactive CSS classes depending on whether the"
Add-Content $ReportFile "// element is active or not."
Add-Content $ReportFile "function setClass(elem, active) {"
Add-Content $ReportFile "   var classes = elem.className.split(/\s+/);"
Add-Content $ReportFile "   var cls = active ? activeClass : inactiveClass, found = false;"
Add-Content $ReportFile "   for (var i = 0; i < classes.length; i++) {"
Add-Content $ReportFile "       if (classes[i] == activeClass || classes[i] == inactiveClass) {"
Add-Content $ReportFile "           if (!found) {"
Add-Content $ReportFile "               classes[i] = cls;"
Add-Content $ReportFile "               found = true;"
Add-Content $ReportFile "           } else {"
Add-Content $ReportFile "               delete classes[i--];"
Add-Content $ReportFile "           }"
Add-Content $ReportFile "       }"
Add-Content $ReportFile "   }"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   if (!found) classes.push(cls);"
Add-Content $ReportFile "   elem.className = classes.join(' ');"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "// Functions for managing the hash."
Add-Content $ReportFile "function getParams() {"
Add-Content $ReportFile "   var hash = location.hash || '#';"
Add-Content $ReportFile "   var parts = hash.substring(1).split('&');"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   var params = {};"
Add-Content $ReportFile "   for (var i = 0; i < parts.length; i++) {"
Add-Content $ReportFile "       var nv = parts[i].split('=');"
Add-Content $ReportFile "       if (!nv[0]) continue;"
Add-Content $ReportFile "       params[nv[0]] = nv[1] || null;"
Add-Content $ReportFile "   }"
Add-Content $ReportFile "   "   
Add-Content $ReportFile "   return params;"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "function setParams(params) {"
Add-Content $ReportFile "   var parts = [];"
Add-Content $ReportFile "   for (var name in params) {"
Add-Content $ReportFile "       // One of the following two lines of code must be commented out. Use the"
Add-Content $ReportFile "       // first to keep empty values in the hash query string; use the second"
Add-Content $ReportFile "       // to remove them."
Add-Content $ReportFile "       //parts.push(params[name] ? name + '=' + params[name] : name);"
Add-Content $ReportFile "       if (params[name]) parts.push(name + '=' + params[name]);"
Add-Content $ReportFile "   }"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   location.hash = knownHash = '#' + parts.join('&');"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "// Looks for changes to the hash."
Add-Content $ReportFile "var knownHash = location.hash;"
Add-Content $ReportFile "function pollHash() {"
Add-Content $ReportFile "   var hash = location.hash;"
Add-Content $ReportFile "   if (hash != knownHash) {"
Add-Content $ReportFile "       var params = getParams();"
Add-Content $ReportFile "       for (var name in params) {"
Add-Content $ReportFile "           if (!(name in activates)) continue;"
Add-Content $ReportFile "           activates[name](params[name]);"
Add-Content $ReportFile "       }"
Add-Content $ReportFile "       knownHash = hash;"
Add-Content $ReportFile "   }"
Add-Content $ReportFile "}"
Add-Content $ReportFile "setInterval(pollHash, 250);"
Add-Content $ReportFile "   "
Add-Content $ReportFile "function getParam(name) {"
Add-Content $ReportFile "   var params = getParams();"
Add-Content $ReportFile "   return params[name];"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "function setParam(name, value) {"
Add-Content $ReportFile "   var params = getParams();"
Add-Content $ReportFile "   params[name] = value;"
Add-Content $ReportFile "   setParams(params);"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "// If the hash is currently set to something that looks like a single id,"
Add-Content $ReportFile "// automatically activate any elements with that id."
Add-Content $ReportFile "var initialId = null;"
Add-Content $ReportFile "var match = regex.exec(knownHash);"
Add-Content $ReportFile "if (match) {"
Add-Content $ReportFile "   initialId = match[1];"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "// Takes an array of either element IDs or a hash with the element ID as the key"
Add-Content $ReportFile "// and an array of sub-element IDs as the value."
Add-Content $ReportFile "// When activating these sub-elements, all parent elements will also be"
Add-Content $ReportFile "// activated in the process."
Add-Content $ReportFile "function makeActivatable(paramName, activatables) {"
Add-Content $ReportFile "   var all = {}, first = initialId;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   // Activates all elements for a specific id (and inactivates the others.)"
Add-Content $ReportFile "   function activate(id) {"
Add-Content $ReportFile "       if (!(id in all)) return false;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "       for (var cur in all) {"
Add-Content $ReportFile "           if (cur == id) continue;"
Add-Content $ReportFile "           for (var i = 0; i < all[cur].length; i++) {"
Add-Content $ReportFile "               setClass(all[cur][i], false);"
Add-Content $ReportFile "           }"
Add-Content $ReportFile "       }"
Add-Content $ReportFile "   "
Add-Content $ReportFile "       for (var i = 0; i < all[id].length; i++) {"
Add-Content $ReportFile "           setClass(all[id][i], true);"
Add-Content $ReportFile "       }"
Add-Content $ReportFile "   "
Add-Content $ReportFile "       setParam(paramName, id);"
Add-Content $ReportFile "   "
Add-Content $ReportFile "       return true;"
Add-Content $ReportFile "   }"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   activates[paramName] = activate;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   function attach(item, basePath) {"
Add-Content $ReportFile "       if (item instanceof Array) {"
Add-Content $ReportFile "           for (var i = 0; i < item.length; i++) {"
Add-Content $ReportFile "               attach(item[i], basePath);"
Add-Content $ReportFile "           }"
Add-Content $ReportFile "       } else if (typeof item == 'object') {"
Add-Content $ReportFile "           for (var p in item) {"
Add-Content $ReportFile "               var path = attach(p, basePath);"
Add-Content $ReportFile "               attach(item[p], path);"
Add-Content $ReportFile "           }"
Add-Content $ReportFile "       } else if (typeof item == 'string') {"
Add-Content $ReportFile "           var path = basePath ? basePath.slice(0) : [];"
Add-Content $ReportFile "           var e = document.getElementById(item);"
Add-Content $ReportFile "           if (e)"
Add-Content $ReportFile "               path.push(e);"
Add-Content $ReportFile "           else "
Add-Content $ReportFile "               return;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "           if (!first) first = item;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "           // Store the elements in a lookup table."
Add-Content $ReportFile "           all[item] = path;"
Add-Content $ReportFile "   "
Add-Content $ReportFile "           // Attach a function that will activate the appropriate element"
Add-Content $ReportFile "           // to all anchors."
Add-Content $ReportFile "           if (item in anchors) {"
Add-Content $ReportFile "               // Create a function that will call the 'activate' function with"
Add-Content $ReportFile "               // the proper parameters. It will be used as the event callback."
Add-Content $ReportFile "               var func = (function (id) {"
Add-Content $ReportFile "                   return function (e) {"
Add-Content $ReportFile "                       activate(id);"
Add-Content $ReportFile "   "
Add-Content $ReportFile "                       if (!e) e = window.event;"
Add-Content $ReportFile "                       if (e.preventDefault) e.preventDefault();"
Add-Content $ReportFile "                       e.returnValue = false;"
Add-Content $ReportFile "                       return false;"
Add-Content $ReportFile "                   };"
Add-Content $ReportFile "               })(item);"
Add-Content $ReportFile "   "
Add-Content $ReportFile "               for (var i = 0; i < anchors[item].length; i++) {"
Add-Content $ReportFile "                   var a = anchors[item][i];"
Add-Content $ReportFile "   "
Add-Content $ReportFile "                   if (a.addEventListener) {"
Add-Content $ReportFile "                       a.addEventListener('click', func, false);"
Add-Content $ReportFile "                   } else if (a.attachEvent) {"
Add-Content $ReportFile "                       a.attachEvent('onclick', func);"
Add-Content $ReportFile "                   } else {"
Add-Content $ReportFile "                       throw 'Unsupported event model.';"
Add-Content $ReportFile "                   }"
Add-Content $ReportFile "   "
Add-Content $ReportFile "                   all[item].push(a);"
Add-Content $ReportFile "               }"
Add-Content $ReportFile "           }"
Add-Content $ReportFile "   "
Add-Content $ReportFile "           return path;"
Add-Content $ReportFile "       } else {"
Add-Content $ReportFile "           throw 'Unexpected type.';"
Add-Content $ReportFile "       }"
Add-Content $ReportFile "   "
Add-Content $ReportFile "       return basePath;"
Add-Content $ReportFile "   }"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   attach(activatables);"
Add-Content $ReportFile "   "
Add-Content $ReportFile "   // Activate an element."
Add-Content $ReportFile "   if (first) activate(getParam(paramName)) || activate(first);"
Add-Content $ReportFile "}"
Add-Content $ReportFile "   "
Add-Content $ReportFile "return makeActivatable;"
Add-Content $ReportFile "})();"
Add-Content $ReportFile "   "
Add-Content $ReportFile "activatables('page', $stArray);"
Add-Content $ReportFile "</script>"
Add-Content $ReportFile "</html>"
