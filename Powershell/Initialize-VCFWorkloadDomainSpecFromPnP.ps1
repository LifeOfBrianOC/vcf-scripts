Param(
    [Parameter(Mandatory=$true)][String]$Workbook,
    [Parameter(Mandatory=$true)][String]$Json,
    [Parameter(Mandatory=$true)][String]$nsxtPassword,
    [Parameter(Mandatory=$true)][String]$vCenterPassword
)

Try {
    Write-Host " Importing ImportExcel Module"
    Import-Module ImportExcel -WarningAction SilentlyContinue -ErrorAction Stop
}
Catch {
    Write-Host " ImportExcel Module not found. Installing"
    Install-Module ImportExcel
}
​
Write-Host "Generating Workload Domain JSON"
$pnpWorkbook = Open-ExcelPackage -Path $Workbook
$cidr = ($pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_cidr"].Value.split("/"))[1]
$managmentMaskObject = ([IPAddress] ([Convert]::ToUInt64((("1" * $cidr) + ("0" * (32 - $cidr))), 2)))
​
$nsxtNode1Object = @()
    $nsxtNode1Object += [pscustomobject]@{
        'ipAddress' = $pnpWorkbook.Workbook.Names["wld_nsxt_mgra_ip"].Value
        'dnsName' = $pnpWorkbook.Workbook.Names["wld_nsxt_mgra_fqdn"].Value
        'gateway' = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_gateway_ip"].Value
        'subnetMask' = $managmentMaskObject.IPAddressToString
    }
​
$nsxtNode2Object = @()
    $nsxtNode2Object += [pscustomobject]@{
        'ipAddress' = $pnpWorkbook.Workbook.Names["wld_nsxt_mgrb_ip"].Value
        'dnsName' = $pnpWorkbook.Workbook.Names["wld_nsxt_mgrb_fqdn"].Value
        'gateway' = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_gateway_ip"].Value
        'subnetMask' = $managmentMaskObject.IPAddressToString
    }
​
$nsxtNode3Object = @()
    $nsxtNode3Object += [pscustomobject]@{
        'ipAddress' = $pnpWorkbook.Workbook.Names["wld_nsxt_mgrc_ip"].Value
        'dnsName' = $pnpWorkbook.Workbook.Names["wld_nsxt_mgrc_fqdn"].Value
        'gateway' = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_gateway_ip"].Value
        'subnetMask' = $managmentMaskObject.IPAddressToString
    }
​
$nsxtManagerObject = @()
    $nsxtManagerObject += [pscustomobject]@{
        'name' = $pnpWorkbook.Workbook.Names["wld_nsxt_mgra_hostname"].Value
        networkDetailsSpec = ($nsxtNode1Object | Select-Object -Skip 0)
    }
    $nsxtManagerObject += [pscustomobject]@{
        'name' = $pnpWorkbook.Workbook.Names["wld_nsxt_mgrb_hostname"].Value
        networkDetailsSpec = ($nsxtNode2Object | Select-Object -Skip 0)
    }
    $nsxtManagerObject += [pscustomobject]@{
        'name' = $pnpWorkbook.Workbook.Names["wld_nsxt_mgrc_hostname"].Value
        networkDetailsSpec = ($nsxtNode3Object | Select-Object -Skip 0)
    }
​
$nsxtObject = @()
    $nsxtObject += [pscustomobject]@{
        nsxManagerSpecs = $nsxtManagerObject
        'vip' = $pnpWorkbook.Workbook.Names["wld_nsxt_vip_ip"].Value
        'vipFqdn' = $pnpWorkbook.Workbook.Names["wld_nsxt_vip_fqdn"].Value
        'licenseKey' = $pnpWorkbook.Workbook.Names["nsxt_license"].Value
        'nsxManagerAdminPassword' = $nsxtPassword
    }
​
$vmnicObject = @()
    $vmnicObject += [pscustomobject]@{
        'id' = $pnpWorkbook.Workbook.Names["wld_vss_mgmt_nic"].Value
        'vdsName' = $pnpWorkbook.Workbook.Names["wld_vds_name"].Value
    }
    $vmnicObject += [pscustomobject]@{
        'id' = $pnpWorkbook.Workbook.Names["wld_vds_mgmt_nic"].Value
        'vdsName' = $pnpWorkbook.Workbook.Names["wld_vds_name"].Value
    }
​
$hostnetworkObject = @()
    $hostnetworkObject += [pscustomobject]@{
        vmNics = $vmnicObject
    }
​
$hostObject = @()
    $hostObject += [pscustomobject]@{
        'id' = $pnpWorkbook.Workbook.Names["wld_az1_host1_fqdn"].Value
        'licenseKey' = $pnpWorkbook.Workbook.Names["esx_std_license"].Value
        hostNetworkSpec = ($hostnetworkObject | Select-Object -Skip 0)
    }
    $hostObject += [pscustomobject]@{
        'id' = $pnpWorkbook.Workbook.Names["wld_az1_host2_fqdn"].Value
        'licenseKey' = $pnpWorkbook.Workbook.Names["esx_std_license"].Value
        hostNetworkSpec = ($hostnetworkObject | Select-Object -Skip 0)
    }
    $hostObject += [pscustomobject]@{
        'id' = $pnpWorkbook.Workbook.Names["wld_az1_host3_fqdn"].Value
        'licenseKey' = $pnpWorkbook.Workbook.Names["esx_std_license"].Value
        hostNetworkSpec = ($hostnetworkObject | Select-Object -Skip 0)
    }
    $hostObject += [pscustomobject]@{
        'id' = $pnpWorkbook.Workbook.Names["wld_az1_host4_fqdn"].Value
        'licenseKey' = $pnpWorkbook.Workbook.Names["esx_std_license"].Value
        hostNetworkSpec = ($hostnetworkObject | Select-Object -Skip 0)
    }
​
$portgroupObject = @()
    $portgroupObject += [pscustomobject]@{
        'name' = $pnpWorkbook.Workbook.Names["wld_az1_mgmt_pg"].Value
        'transportType' = "MANAGEMENT"
    }
    $portgroupObject += [pscustomobject]@{
        'name' = $pnpWorkbook.Workbook.Names["wld_az1_vmotion_pg"].Value
        'transportType' = "VMOTION"
    }
    $portgroupObject += [pscustomobject]@{
        'name' = $pnpWorkbook.Workbook.Names["wld_az1_principal_storage_pg"].Value
        'transportType' = "VSAN"
    }
​
$vdsObject = @()
    $vdsObject += [pscustomobject]@{
        'name' = $pnpWorkbook.Workbook.Names["wld_vds_name"].Value
        portGroupSpecs = $portgroupObject
    }
​
$nsxTClusterObject = @()
    $nsxTClusterObject += [pscustomobject]@{
        'geneveVlanId' = [String]($pnpWorkbook.Workbook.Names["wld_host_overlay_vlan"].Value)
    }
​
$nsxClusterObject = @()
    $nsxClusterObject += [pscustomobject]@{
        nsxTClusterSpec = ($nsxTClusterObject | Select-Object -Skip 0)
    }
​
$networkObject = @()
    $networkObject += [pscustomobject]@{
        vdsSpecs = $vdsObject
        nsxClusterSpec = ($nsxClusterObject | Select-Object -Skip 0)
    }
​
$vsanDatastoreObject = @()
    $vsanDatastoreObject += [pscustomobject]@{
        'failuresToTolerate' = "1"
        'licenseKey' = $pnpWorkbook.Workbook.Names["vsan_license"].Value
        'datastoreName' = $pnpWorkbook.Workbook.Names["wld_vsan_datastore"].Value
    }
​
$vsanObject = @()
    $vsanObject += [pscustomobject]@{
        vsanDatastoreSpec = ($vsanDatastoreObject | Select-Object -Skip 0)
    }
​
$clusterObject = @()
    $clusterObject += [pscustomobject]@{
        'name' = $pnpWorkbook.Workbook.Names["wld_cluster"].Value
        hostSpecs = $hostObject
        datastoreSpec = ($vsanObject | Select-Object -Skip 0)
        networkSpec = ($networkObject | Select-Object -Skip 0)
    }
​
$computeObject = @()
    $computeObject += [pscustomobject]@{
        clusterSpecs = $clusterObject
    }
​
$vcenterNetworkObject = @()
    $vcenterNetworkObject += [pscustomobject]@{
        'ipAddress' = $pnpWorkbook.Workbook.Names["wld_vc_ip"].Value
        'dnsName' = $pnpWorkbook.Workbook.Names["wld_vc_fqdn"].Value
        'gateway'= $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_gateway_ip"].Value
        'subnetMask' = $managmentMaskObject.IPAddressToString
    }
​
$vcenterObject = @()
    $vcenterObject += [pscustomobject]@{
        'name' = $pnpWorkbook.Workbook.Names["wld_vc_hostname"].Value
        networkDetailsSpec = ($vcenterNetworkObject | Select-Object -Skip 0)
        'rootPassword' = $vCenterPassword
        'datacenterName' = $pnpWorkbook.Workbook.Names["wld_datacenter"].Value
    }
​
$workloadDomainObject = @()
    $workloadDomainObject += [pscustomobject]@{
        'domainName' = $pnpWorkbook.Workbook.Names["wld_sddc_domain"].Value
        'orgName' = $pnpWorkbook.Workbook.Names["wld_sddc_org"].Value
        vcenterSpec = ($vcenterObject | Select-Object -Skip 0)
        computeSpec = ($computeObject | Select-Object -Skip 0)
        nsxTSpec = ($nsxtObject | Select-Object -Skip 0)
    }
​
$workloadDomainObject | ConvertTo-Json -Depth 11 | Out-File -FilePath $Json
Close-ExcelPackage $pnpWorkbook -ErrorAction SilentlyContinue
