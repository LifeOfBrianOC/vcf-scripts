Function LogMessage {
    Param (
        [Parameter (Mandatory = $true)] [AllowEmptyString()] [String]$message,
        [Parameter (Mandatory = $false)] [ValidateSet("INFO", "ERROR", "WARNING", "EXCEPTION", "ADVISORY", "NOTE", "QUESTION", "WAIT")] [String]$type = "INFO",
        [Parameter (Mandatory = $false)] [String]$colour,
        [Parameter (Mandatory = $false)] [Switch]$skipnewline
    )

    If (!$colour) {
        $colour = "92m" #Green
    }

    If ($type -eq "INFO") {
        $messageColour = "92m" #Green
    }
    elseIf ($type -in "ERROR", "EXCEPTION") {
        $messageColour = "91m" # Red
    }
    elseIf ($type -in "WARNING", "ADVISORY", "QUESTION") {
        $messageColour = "93m" #Yellow
    }
    elseIf ($type -in "NOTE", "WAIT") {
        $messageColour = "97m" # White
    }

    If (!$threadTag) { $threadTag = "..."; $threadColour = "97m" }

    <#
    Reference Colours
    31m Red
    32m Green
    33m Yellow
    36m Cyan
    37m White
    91m Bright Red
    92m Bright Green
    93m Bright Yellow
    95m Bright Magenta
    96m Bright Cyan
    97m Bright White
    #>

    $ESC = [char]0x1b
    $timestampColour = "97m"

    $timeStamp = Get-Date -Format "MM-dd-yyyy_HH:mm:ss"

    $threadTag = $threadTag.toUpper()
    If ($headlessPassed) {
        If ($skipnewline) {
            Write-Host -NoNewline "[$timestamp] [$threadTag] [$type] $message"
        }
        else {
            Write-Host "[$timestamp] [$threadTag] [$type] $message"
        }
    }
    else {
        If ($skipnewline) {
            Write-Host -NoNewline "$ESC[${timestampcolour} [$timestamp]$ESC[${threadColour} [$threadTag]$ESC[${messageColour} [$type] $message$ESC[0m"
        }
        else {
            Write-Host "$ESC[${timestampcolour} [$timestamp]$ESC[${threadColour} [$threadTag]$ESC[${messageColour} [$type] $message$ESC[0m"
        }
    }
    $logContent = '[' + $timeStamp + '] [' + $threadTag + '] ' + $type + ' ' + $message
    Add-Content -path $logFile $logContent
}


Function catchWriter {
    Param (
        [Parameter (mandatory = $true)] [PSObject]$object
    )
    $lineNumber = $object.InvocationInfo.ScriptLineNumber
    $lineText = $object.InvocationInfo.Line.trim()
    $errorMessage = $object.Exception.Message
    LogMessage -Type EXCEPTION -Message "Error at Script Line $lineNumber"
    LogMessage -Type EXCEPTION -Message "Relevant Command: $lineText"
    LogMessage -Type EXCEPTION -Message "Error Message: $errorMessage"
}


Function Initialize-VCFManagementDomainSpecFromPnP {

    Param (
        [Parameter (Mandatory = $true)] [String]$workbook,
        [Parameter (Mandatory = $true)] [String]$path
    )
    $Global:vcfVersion = @("v4.3.x", "v4.4.x", "v4.5.x", "v5.0.x", "v5.1.x")
    Try {
        
        $module = "Management Domain JSON Spec"
        Write-Host "Starting the Process of Generating the $module"
        $pnpWorkbook = Open-ExcelPackage -Path $Workbook

        If ($pnpWorkbook.Workbook.Names["vcf_version"].Value -notin $vcfVersion) {
            Write-Host "Planning and Preparation Workbook Provided Not Supported"
            Break
        }

        If ($pnpWorkbook.Workbook.Names["vcf_plus_result"].Value -eq "Included") {
            $nsxtLicense = ""
            $esxLicense = ""
            $vsanLicense = ""
            $vcenterLicense = ""
        }
        else {
            $nsxtLicense = $pnpWorkbook.Workbook.Names["nsxt_license"].Value
            $esxLicense = $pnpWorkbook.Workbook.Names["esx_std_license"].Value
            $vsanLicense = $pnpWorkbook.Workbook.Names["vsan_license"].Value
            $vcenterLicense = $pnpWorkbook.Workbook.Names["vc_license"].Value
        }

        Write-Host "Generating the $module"

        $cidr = ($pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_cidr"].Value.split("/"))[1]
        $managmentMaskObject = ([IPAddress] ([Convert]::ToUInt64((("1" * $cidr) + ("0" * (32 - $cidr))), 2)))

        $ntpServers = New-Object System.Collections.ArrayList
        If ($pnpWorkbook.Workbook.Names["region_dns2_ip"].Value -eq "n/a") {
            [Array]$ntpServers = $pnpWorkbook.Workbook.Names["region_dns1_ip"].Value
        }
        else {
            [Array]$ntpServers = $pnpWorkbook.Workbook.Names["region_dns1_ip"].Value, $pnpWorkbook.Workbook.Names["region_dns2_ip"].Value
        }

        $dnsObject = @()
        $dnsObject += [pscustomobject]@{
            'domain'              = $pnpWorkbook.Workbook.Names["region_ad_parent_fqdn"].Value
            'subdomain'           = $pnpWorkbook.Workbook.Names["region_ad_child_fqdn"].Value
            'nameserver'          = $pnpWorkbook.Workbook.Names["region_dns1_ip"].Value
            'secondaryNameserver' = $pnpWorkbook.Workbook.Names["region_dns2_ip"].Value
        }

        $rootUserObject = @()
        $rootUserObject += [pscustomobject]@{
            'username' = "root"
            'password' = $pnpWorkbook.Workbook.Names["sddc_mgr_root_password"].Value
        }

        $secondUserObject = @()
        $secondUserObject += [pscustomobject]@{
            'username' = "vcf"
            'password' = $pnpWorkbook.Workbook.Names["sddc_mgr_vcf_password"].Value
        }

        $restApiUserObject = @()
        $restApiUserObject += [pscustomobject]@{
            'username' = "admin"
            'password' = $pnpWorkbook.Workbook.Names["sddc_mgr_admin_local_password"].Value
        }

        $sddcManagerObject = @()
        $sddcManagerObject += [pscustomobject]@{
            'hostname'            = $pnpWorkbook.Workbook.Names["sddc_mgr_hostname"].Value
            'ipAddress'           = $pnpWorkbook.Workbook.Names["sddc_mgr_ip"].Value
            'netmask'             = $managmentMaskObject.IPAddressToString
            'localUserPassword'   = $pnpWorkbook.Workbook.Names["sddc_mgr_admin_local_password"].Value
            rootUserCredentials   = ($rootUserObject | Select-Object -Skip 0)
            restApiCredentials    = ($restApiUserObject | Select-Object -Skip 0)
            secondUserCredentials = ($secondUserObject | Select-Object -Skip 0)
        }

        $vmnics = New-Object System.Collections.ArrayList
        [Array]$vmnics = $($pnpWorkbook.Workbook.Names["primary_vds_vmnics"].Value.Split(',')[0]), $($pnpWorkbook.Workbook.Names["primary_vds_vmnics"].Value.Split(',')[1])

        $networks = New-Object System.Collections.ArrayList
        # Commented out as breaks 4.5
        [Array]$networks = "MANAGEMENT", "VMOTION", "VSAN"

        $vmotionIpObject = @()
        $vmotionIpObject += [pscustomobject]@{
            'startIpAddress' = $pnpWorkbook.Workbook.Names["mgmt_az1_vmotion_pool_start_ip"].Value
            'endIpAddress'   = $pnpWorkbook.Workbook.Names["mgmt_az1_vmotion_pool_end_ip"].Value
        }

        $vsanIpObject = @()
        $vsanIpObject += [pscustomobject]@{
            'startIpAddress' = $pnpWorkbook.Workbook.Names["mgmt_az1_vsan_pool_start_ip"].Value
            'endIpAddress'   = $pnpWorkbook.Workbook.Names["mgmt_az1_vsan_pool_end_ip"].Value
        }

        $vmotionMtu = $pnpWorkbook.Workbook.Names["mgmt_az1_vmotion_mtu"].Value -as [string]
        $vsanMtu = $pnpWorkbook.Workbook.Names["mgmt_az1_vsan_mtu"].Value -as [string]
        $dvsMtu = [INT]$pnpWorkbook.Workbook.Names["primary_vds_mtu"].Value

        $networkObject = @()
        $networkObject += [pscustomobject]@{
            'networkType'  = "MANAGEMENT"
            'subnet'       = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_cidr"].Value
            'vlanId'       = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_vlan"].Value -as [string]
            'mtu'          = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_mtu"].Value -as [string]
            'gateway'      = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_gateway_ip"].Value
            'portGroupKey' = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_pg"].Value
        }
        $networkObject += [pscustomobject]@{
            'networkType'          = "VMOTION"
            'subnet'               = $pnpWorkbook.Workbook.Names["mgmt_az1_vmotion_cidr"].Value
            includeIpAddressRanges = $vmotionIpObject
            'vlanId'               = $pnpWorkbook.Workbook.Names["mgmt_az1_vmotion_vlan"].Value -as [string]
            'mtu'                  = $pnpWorkbook.Workbook.Names["mgmt_az1_vmotion_mtu"].Value -as [string]
            'gateway'              = $pnpWorkbook.Workbook.Names["mgmt_az1_vmotion_gateway_ip"].Value
            'portGroupKey'         = $pnpWorkbook.Workbook.Names["mgmt_az1_vmotion_pg"].Value
        }
        $networkObject += [pscustomobject]@{
            'networkType'          = "VSAN"
            'subnet'               = $pnpWorkbook.Workbook.Names["mgmt_az1_vsan_cidr"].Value
            includeIpAddressRanges = $vsanIpObject
            'vlanId'               = $pnpWorkbook.Workbook.Names["mgmt_az1_vsan_vlan"].Value -as [string]
            'mtu'                  = $pnpWorkbook.Workbook.Names["mgmt_az1_vsan_mtu"].Value -as [string]
            'gateway'              = $pnpWorkbook.Workbook.Names["mgmt_az1_vsan_gateway_ip"].Value
            'portGroupKey'         = $pnpWorkbook.Workbook.Names["mgmt_az1_vsan_pg"].Value
        }
        
        $nsxtManagerObject = @()
        $nsxtManagerObject += [pscustomobject]@{
            'hostname' = $pnpWorkbook.Workbook.Names["mgmt_nsxt_mgra_hostname"].Value
            'ip'       = $pnpWorkbook.Workbook.Names["mgmt_nsxt_mgra_ip"].Value
        }
        If ($singleNSXTManager -eq "N") {
            $nsxtManagerObject += [pscustomobject]@{
                'hostname' = $pnpWorkbook.Workbook.Names["mgmt_nsxt_mgrb_hostname"].Value
                'ip'       = $pnpWorkbook.Workbook.Names["mgmt_nsxt_mgrb_ip"].Value
            }
            $nsxtManagerObject += [pscustomobject]@{
                'hostname' = $pnpWorkbook.Workbook.Names["mgmt_nsxt_mgrc_hostname"].Value
                'ip'       = $pnpWorkbook.Workbook.Names["mgmt_nsxt_mgrc_ip"].Value
            }
        }

        $vlanTransportZoneObject = @()
        $vlanTransportZoneObject += [pscustomobject]@{
            'zoneName'    = $pnpWorkbook.Workbook.Names["mgmt_sddc_domain"].Value + "-tz-vlan01"
            'networkName' = "netName-vlan"
        }

        $overlayTransportZoneObject = @()
        $overlayTransportZoneObject += [pscustomobject]@{
            'zoneName'    = $pnpWorkbook.Workbook.Names["mgmt_sddc_domain"].Value + "-tz-overlay01"
            'networkName' = "netName-overlay"
        }

        $edgeNode01interfaces = @()
        $edgeNode01interfaces += [pscustomobject]@{
            'name'          = $pnpWorkbook.Workbook.Names["mgmt_sddc_domain"].Value + "-uplink01-tor1"
            'interfaceCidr' = $pnpWorkbook.Workbook.Names["mgmt_en1_edge_overlay_interface_ip_1_ip"].Value
        }
        $edgeNode01interfaces += [pscustomobject]@{
            'name'          = $pnpWorkbook.Workbook.Names["mgmt_sddc_domain"].Value + "-uplink01-tor2"
            'interfaceCidr' = $pnpWorkbook.Workbook.Names["mgmt_en1_edge_overlay_interface_ip_2_ip"].Value
        }

        $edgeNode02interfaces = @()
        $edgeNode02interfaces += [pscustomobject]@{
            'name'          = $pnpWorkbook.Workbook.Names["mgmt_sddc_domain"].Value + "-uplink01-tor1"
            'interfaceCidr' = $pnpWorkbook.Workbook.Names["mgmt_en2_edge_overlay_interface_ip_1_ip"].Value
        }
        $edgeNode02interfaces += [pscustomobject]@{
            'name'          = $pnpWorkbook.Workbook.Names["mgmt_sddc_domain"].Value + "-uplink01-tor2"
            'interfaceCidr' = $pnpWorkbook.Workbook.Names["mgmt_en2_edge_overlay_interface_ip_2_ip"].Value

        }
        
        $edgeNodeObject = @()
        $edgeNodeObject += [pscustomobject]@{
            'edgeNodeName'     = $pnpWorkbook.Workbook.Names["mgmt_en1_fqdn"].Value.Split(".")[0]
            'edgeNodeHostname' = $pnpWorkbook.Workbook.Names["mgmt_en1_fqdn"].Value
            'managementCidr'   = $pnpWorkbook.Workbook.Names["input_mgmt_en1_ip"].Value + "/" + $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_cidr"].Value.Split("/")[-1]
            'edgeVtep1Cidr'    = $pnpWorkbook.Workbook.Names["input_mgmt_en1_edge_overlay_interface_ip_1_ip"].Value + "/" + $pnpWorkbook.Workbook.Names["input_mgmt_edge_overlay_cidr"].Value.Split("/")[-1]
            'edgeVtep2Cidr'    = $pnpWorkbook.Workbook.Names["input_mgmt_en1_edge_overlay_interface_ip_2_ip"].Value + "/" + $pnpWorkbook.Workbook.Names["input_mgmt_edge_overlay_cidr"].Value.Split("/")[-1]
            interfaces         = $edgeNode01interfaces
        }        
        $edgeNodeObject += [pscustomobject]@{
            'edgeNodeName'     = $pnpWorkbook.Workbook.Names["mgmt_en2_fqdn"].Value.Split(".")[0]
            'edgeNodeHostname' = $pnpWorkbook.Workbook.Names["mgmt_en2_fqdn"].Value
            'managementCidr'   = $pnpWorkbook.Workbook.Names["input_mgmt_en2_ip"].Value + "/" + $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_cidr"].Value.Split("/")[-1]
            'edgeVtep1Cidr'    = $pnpWorkbook.Workbook.Names["input_mgmt_en2_edge_overlay_interface_ip_1_ip"].Value + "/" + $pnpWorkbook.Workbook.Names["input_mgmt_edge_overlay_cidr"].Value.Split("/")[-1]
            'edgeVtep2Cidr'    = $pnpWorkbook.Workbook.Names["input_mgmt_en2_edge_overlay_interface_ip_2_ip"].Value + "/" + $pnpWorkbook.Workbook.Names["input_mgmt_edge_overlay_cidr"].Value.Split("/")[-1]
            interfaces         = $edgeNode02interfaces
        }
        
        $edgeServicesObject = @()
        $edgeServicesObject += [pscustomobject]@{
            'tier0GatewayName' = $pnpWorkbook.Workbook.Names["mgmt_tier0_name"].Value
            'tier1GatewayName' = $pnpWorkbook.Workbook.Names["mgmt_tier1_name"].Value
        }

        $bgpNeighboursObject = @()
        $bgpNeighboursObject += [pscustomobject]@{
            'neighbourIp'      = $pnpWorkbook.Workbook.Names["input_mgmt_az1_tor1_peer_ip"].Value
            'autonomousSystem' = $pnpWorkbook.Workbook.Names["input_mgmt_az1_tor1_peer_asn"].Value
            'password'         = $pnpWorkbook.Workbook.Names["input_mgmt_az1_tor1_peer_bgp_password"].Value
        }
        $bgpNeighboursObject += [pscustomobject]@{
            'neighbourIp'      = $pnpWorkbook.Workbook.Names["input_mgmt_az1_tor2_peer_ip"].Value
            'autonomousSystem' = $pnpWorkbook.Workbook.Names["input_mgmt_az1_tor2_peer_asn"].Value
            'password'         = $pnpWorkbook.Workbook.Names["input_mgmt_az1_tor2_peer_bgp_password"].Value
        }

        $nsxtEdgeObject = @()
        $nsxtEdgeObject += [pscustomobject]@{
            'edgeClusterName'               = $pnpWorkbook.Workbook.Names["mgmt_ec_name"].Value
            'edgeRootPassword'              = $pnpWorkbook.Workbook.Names["nsxt_en_root_password"].Value
            'edgeAdminPassword'             = $pnpWorkbook.Workbook.Names["nsxt_en_admin_password"].Value
            'edgeAuditPassword'             = $pnpWorkbook.Workbook.Names["nsxt_en_audit_password"].Value
            'edgeFormFactor'                = $pnpWorkbook.Workbook.Names["mgmt_ec_formfactor"].Value 
            'tier0ServicesHighAvailability' = "ACTIVE_ACTIVE"
            'asn'                           = $pnpWorkbook.Workbook.Names["mgmt_en_asn"].Value
            edgeServicesSpecs               = ($edgeServicesObject | Select-Object -Skip 0)
            edgeNodeSpecs                   = $edgeNodeObject
            bgpNeighbours                   = $bgpNeighboursObject
        }
        
        $logicalSegmentsObject = @()
        $logicalSegmentsObject += [pscustomobject]@{
            'name'        = $pnpWorkbook.Workbook.Names["reg_seg01_name"].Value
            'networkType' = "REGION_SPECIFIC"
        }
        $logicalSegmentsObject += [pscustomobject]@{
            'name'        = $pnpWorkbook.Workbook.Names["xreg_seg01_name"].Value
            'networkType' = "X_REGION"
        }

        $nsxtObject = @()
        $nsxtObject += [pscustomobject]@{
            'nsxtManagerSize'                = $pnpWorkbook.Workbook.Names["mgmt_nsxt_mgr_formfactor"].Value.tolower()
            nsxtManagers                     = $nsxtManagerObject
            'rootNsxtManagerPassword'        = $pnpWorkbook.Workbook.Names["nsxt_lm_root_password"].Value
            'nsxtAdminPassword'              = $pnpWorkbook.Workbook.Names["nsxt_lm_admin_password"].Value
            'nsxtAuditPassword'              = $pnpWorkbook.Workbook.Names["nsxt_lm_audit_password"].Value
            'rootLoginEnabledForNsxtManager' = "true"
            'sshEnabledForNsxtManager'       = "true"
            overLayTransportZone             = ($overlayTransportZoneObject | Select-Object -Skip 0)
            vlanTransportZone                = ($vlanTransportZoneObject | Select-Object -Skip 0)
            'vip'                            = $pnpWorkbook.Workbook.Names["mgmt_nsxt_vip_ip"].Value
            'vipFqdn'                        = $pnpWorkbook.Workbook.Names["mgmt_nsxt_hostname"].Value
            'nsxtLicense'                    = $nsxtLicense
            'transportVlanId'                = $pnpWorkbook.Workbook.Names["mgmt_az1_host_overlay_vlan"].Value -as [string]
        }

        $excelvsanDedup = $pnpWorkbook.Workbook.Names["mgmt_vsan_dedup"].Value
        If ($excelvsanDedup -eq "No") {
            $vsanDedup = $false
        }
        elseIf ($excelvsanDedup -eq "Yes") {
            $vsanDedup = $true
        }

        If ($pnpWorkbook.Workbook.Names["mgmt_principal_storage_chosen"].Value -eq "vSAN-ESA") {
            $ESAenabledtrueobject = @()
            $ESAenabledtrueobject += [pscustomobject]@{
                'enabled' = "true"
            }
        }
        else {
            $ESAenabledtrueobject = @()
            $ESAenabledtrueobject += [pscustomobject]@{
                'enabled' = "false"
            } 
        }

        $vsanObject = @()
        If ($pnpWorkbook.Workbook.Names["mgmt_principal_storage_chosen"].Value -eq "vSAN-ESA") {
            $vsanObject += [pscustomobject]@{
                'vsanName'      = "vsan-1"
                'licenseFile'   = $vsanLicense
                'vsanDedup'     = $vsanDedup
                'datastoreName' = $pnpWorkbook.Workbook.Names["mgmt_vsan_datastore"].Value
                esaConfig       = ($ESAenabledtrueobject | Select-Object -Skip 0)
            }
        }
        else {
            $vsanObject += [pscustomobject]@{
                'vsanName'      = "vsan-1"
                'licenseFile'   = $vsanLicense
                'vsanDedup'     = $vsanDedup
                'datastoreName' = $pnpWorkbook.Workbook.Names["mgmt_vsan_datastore"].Value
            }
        }
        $niocObject = @()
        $niocObject += [pscustomobject]@{
            'trafficType' = "VSAN"
            'value'       = "HIGH"
        }
        $niocObject += [pscustomobject]@{
            'trafficType' = "VMOTION"
            'value'       = "LOW"
        }
        $niocObject += [pscustomobject]@{
            'trafficType' = "VDP"
            'value'       = "LOW"
        }
        $niocObject += [pscustomobject]@{
            'trafficType' = "VIRTUALMACHINE"
            'value'       = "HIGH"
        }
        $niocObject += [pscustomobject]@{
            'trafficType' = "MANAGEMENT"
            'value'       = "NORMAL"
        }
        $niocObject += [pscustomobject]@{
            'trafficType' = "NFS"
            'value'       = "LOW"
        }
        $niocObject += [pscustomobject]@{
            'trafficType' = "HBR"
            'value'       = "LOW"
        }
        $niocObject += [pscustomobject]@{
            'trafficType' = "FAULTTOLERANCE"
            'value'       = "LOW"
        }
        $niocObject += [pscustomobject]@{
            'trafficType' = "ISCSI"
            'value'       = "LOW"
        }

        $dvsObject = @()
        $dvsObject += [pscustomobject]@{
            'mtu'      = $dvsMtu
            niocSpecs  = $niocObject
            'dvsName'  = $pnpWorkbook.Workbook.Names["primary_vds_name"].Value
            'vmnics'   = $vmnics
            'networks' = $networks
        }

        $vmFolderObject = @()
        $vmFOlderObject += [pscustomobject]@{
            'MANAGEMENT' = $pnpWorkbook.Workbook.Names["mgmt_mgmt_vm_folder"].Value
            'NETWORKING' = $pnpWorkbook.Workbook.Names["mgmt_nsx_vm_folder"].Value
            'EDGENODES'  = $pnpWorkbook.Workbook.Names["mgmt_edge_vm_folder"].Value
        }

        If (($pnpWorkbook.Workbook.Names["mgmt_evc_mode"].Value -eq "n/a") -or ($pnpWorkbook.Workbook.Names["mgmt_evc_mode"].Value -eq $null)) {
            $evcMode = ""
        }
        else {
            $evcMode = $pnpWorkbook.Workbook.Names["mgmt_evc_mode"].Value
        }

        $resourcePoolObject = @()
        $resourcePoolObject += [pscustomobject]@{
            'type'                        = "management"
            'name'                        = $pnpWorkbook.Workbook.Names["mgmt_mgmt_rp"].Value
            'cpuSharesLevel'              = "high"
            'cpuSharesValue'              = "0" -as [int]
            'cpuLimit'                    = "-1" -as [int]
            'cpuReservationExpandable'    = $true
            'cpuReservationPercentage'    = "0" -as [int]
            'memorySharesLevel'           = "normal"
            'memorySharesValue'           = "0" -as [int]
            'memoryLimit'                 = "-1" -as [int]
            'memoryReservationExpandable' = $true
            'memoryReservationPercentage' = "0" -as [int]
        }
        $resourcePoolObject += [pscustomobject]@{
            'type'                        = "network"
            'name'                        = $pnpWorkbook.Workbook.Names["mgmt_nsx_rp"].Value
            'cpuSharesLevel'              = "high"
            'cpuSharesValue'              = "0" -as [int]
            'cpuLimit'                    = "-1" -as [int]
            'cpuReservationExpandable'    = $true
            'cpuReservationPercentage'    = "0" -as [int]
            'memorySharesLevel'           = "normal"
            'memorySharesValue'           = "0" -as [int]
            'memoryLimit'                 = "-1" -as [int]
            'memoryReservationExpandable' = $true
            'memoryReservationPercentage' = "0" -as [int]
        }
        $resourcePoolObject += [pscustomobject]@{
            'type'                        = "compute"
            'name'                        = $pnpWorkbook.Workbook.Names["mgmt_user_edge_rp"].Value
            'cpuSharesLevel'              = "normal"
            'cpuSharesValue'              = "0" -as [int]
            'cpuLimit'                    = "-1" -as [int]
            'cpuReservationExpandable'    = $true
            'cpuReservationPercentage'    = "0" -as [int]
            'memorySharesLevel'           = "normal"
            'memorySharesValue'           = "0" -as [int]
            'memoryLimit'                 = "-1" -as [int]
            'memoryReservationExpandable' = $true
            'memoryReservationPercentage' = "0" -as [int]
        }
        $resourcePoolObject += [pscustomobject]@{
            'type'                        = "compute"
            'name'                        = $pnpWorkbook.Workbook.Names["mgmt_user_vm_rp"].Value
            'cpuSharesLevel'              = "normal"
            'cpuSharesValue'              = "0" -as [int]
            'cpuLimit'                    = "-1" -as [int]
            'cpuReservationExpandable'    = $true
            'cpuReservationPercentage'    = "0" -as [int]
            'memorySharesLevel'           = "normal"
            'memorySharesValue'           = "0" -as [int]
            'memoryLimit'                 = "-1" -as [int]
            'memoryReservationExpandable' = $true
            'memoryReservationPercentage' = "0" -as [int]
        }

        If ($pnpWorkbook.Workbook.Names["mgmt_consolidated_result"].Value -eq "Included") {
            $clusterObject = @()
            $clusterObject += [pscustomobject]@{
                vmFolders         = ($vmFolderObject | Select-Object -Skip 0)
                'clusterName'     = $pnpWorkbook.Workbook.Names["mgmt_cluster"].Value
                'clusterEvcMode'  = $evcMode
                resourcePoolSpecs = $resourcePoolObject
            }
        }
        else {
            $clusterObject = @()
            $clusterObject += [pscustomobject]@{
                vmFolders        = ($vmFolderObject | Select-Object -Skip 0)
                'clusterName'    = $pnpWorkbook.Workbook.Names["mgmt_cluster"].Value
                'clusterEvcMode' = $evcMode
            }
        }

        $ssoObject = @()
        $ssoObject += [pscustomobject]@{
            'ssoDomain' = 'vsphere.local'
        }

        $pscObject = @()
        $pscObject += [pscustomobject]@{
            pscSsoSpec             = ($ssoObject | Select-Object -Skip 0)
            'adminUserSsoPassword' = $pnpWorkbook.Workbook.Names["administrator_vsphere_local_password"].Value
        }

        $vcenterObject = @()
        $vcenterObject += [pscustomobject]@{
            'vcenterIp'           = $pnpWorkbook.Workbook.Names["mgmt_vc_ip"].Value
            'vcenterHostname'     = $pnpWorkbook.Workbook.Names["mgmt_vc_hostname"].Value
            'licenseFile'         = $vcenterLicense
            'rootVcenterPassword' = $pnpWorkbook.Workbook.Names["vcenter_root_password"].Value
            'vmSize'              = $pnpWorkbook.Workbook.Names["mgmt_vc_size"].Value.tolower()
        }

        $hostCredentialsObject = @()
        $hostCredentialsObject += [pscustomobject]@{
            'username' = 'root'
            'password' = $pnpWorkbook.Workbook.Names["esxi_root_password"].Value
        }

        $mgmtHost01Object = @()
        $mgmtHost01Object += [pscustomobject]@{
            'subnet'    = $managmentMaskObject.IPAddressToString
            'ipAddress' = $pnpWorkbook.Workbook.Names["mgmt_az1_host1_mgmt_ip"].Value
            'gateway'   = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_gateway_ip"].Value
        }

        $mgmtHost02Object = @()
        $mgmtHost02Object += [pscustomobject]@{
            'subnet'    = $managmentMaskObject.IPAddressToString
            'ipAddress' = $pnpWorkbook.Workbook.Names["mgmt_az1_host2_mgmt_ip"].Value
            'gateway'   = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_gateway_ip"].Value
        }

        $mgmtHost03Object = @()
        $mgmtHost03Object += [pscustomobject]@{
            'subnet'    = $managmentMaskObject.IPAddressToString
            'ipAddress' = $pnpWorkbook.Workbook.Names["mgmt_az1_host3_mgmt_ip"].Value
            'gateway'   = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_gateway_ip"].Value
        }

        $mgmtHost04Object = @()
        $mgmtHost04Object += [pscustomobject]@{
            'subnet'    = $managmentMaskObject.IPAddressToString
            'ipAddress' = $pnpWorkbook.Workbook.Names["mgmt_az1_host4_mgmt_ip"].Value
            'gateway'   = $pnpWorkbook.Workbook.Names["mgmt_az1_mgmt_gateway_ip"].Value
        }

        $HostObject = @()
        $HostObject += [pscustomobject]@{
            'hostname'    = $pnpWorkbook.Workbook.Names["mgmt_az1_host1_hostname"].Value
            'vSwitch'     = $pnpWorkbook.Workbook.Names["mgmt_vss_switch"].Value
            'association' = $pnpWorkbook.Workbook.Names["mgmt_datacenter"].Value
            credentials   = ($hostCredentialsObject | Select-Object -Skip 0)
            ipAddress     = ($mgmtHost01Object | Select-Object -Skip 0)
        }
        $HostObject += [pscustomobject]@{
            'hostname'    = $pnpWorkbook.Workbook.Names["mgmt_az1_host2_hostname"].Value
            'vSwitch'     = $pnpWorkbook.Workbook.Names["mgmt_vss_switch"].Value
            'association' = $pnpWorkbook.Workbook.Names["mgmt_datacenter"].Value
            credentials   = ($hostCredentialsObject | Select-Object -Skip 0)
            ipAddress     = ($mgmtHost02Object | Select-Object -Skip 0)
        }
        $HostObject += [pscustomobject]@{
            'hostname'    = $pnpWorkbook.Workbook.Names["mgmt_az1_host3_hostname"].Value
            'vSwitch'     = $pnpWorkbook.Workbook.Names["mgmt_vss_switch"].Value
            'association' = $pnpWorkbook.Workbook.Names["mgmt_datacenter"].Value
            credentials   = ($hostCredentialsObject | Select-Object -Skip 0)
            ipAddress     = ($mgmtHost03Object | Select-Object -Skip 0)
        }
        $HostObject += [pscustomobject]@{
            'hostname'    = $pnpWorkbook.Workbook.Names["mgmt_az1_host4_hostname"].Value
            'vSwitch'     = $pnpWorkbook.Workbook.Names["mgmt_vss_switch"].Value
            'association' = $pnpWorkbook.Workbook.Names["mgmt_datacenter"].Value
            credentials   = ($hostCredentialsObject | Select-Object -Skip 0)
            ipAddress     = ($mgmtHost04Object | Select-Object -Skip 0)
        }

        $excluded = New-Object System.Collections.ArrayList
        [Array]$excluded = "NSX-V"

        $ceipState = $pnpWorkbook.Workbook.Names["mgmt_ceip_status"].Value
        If ($ceipState -eq "Yes") {
            $ceipEnabled = "$true"
        }
        else {
            $ceipEnabled = "$false"
        }

        $fipsState = $pnpWorkbook.Workbook.Names["mgmt_fips_status"].Value
        If ($fipsState -eq "Yes") {
            $fipsEnabled = "$true"
        }
        else {
            $fipsEnabled = "$false"
        }
        
        $managementDomainObject = New-Object -TypeName psobject
        $managementDomainObject | Add-Member -notepropertyname 'taskName' -notepropertyvalue "workflowconfig/workflowspec-ems.json"
        $managementDomainObject | Add-Member -notepropertyname 'sddcId' -notepropertyvalue $pnpWorkbook.Workbook.Names["mgmt_sddc_domain"].Value
        $managementDomainObject | Add-Member -notepropertyname 'ceipEnabled' -notepropertyvalue $ceipEnabled
        $managementDomainObject | Add-Member -notepropertyname 'fipsEnabled' -notepropertyvalue $fipsEnabled
        $managementDomainObject | Add-Member -notepropertyname 'managementPoolName' -notepropertyvalue $pnpWorkbook.Workbook.Names["mgmt_az1_pool_name"].Value
        $managementDomainObject | Add-Member -notepropertyname 'skipEsxThumbprintValidation' -notepropertyvalue $true
        $managementDomainObject | Add-Member -notepropertyname 'esxLicense' -notepropertyvalue $esxLicense
        $managementDomainObject | Add-Member -notepropertyname 'excludedComponents' -notepropertyvalue $excluded
        $managementDomainObject | Add-Member -notepropertyname 'ntpServers' -notepropertyvalue $ntpServers
        $managementDomainObject | Add-Member -notepropertyname 'dnsSpec' -notepropertyvalue ($dnsObject | Select-Object -Skip 0)
        $managementDomainObject | Add-Member -notepropertyname 'sddcManagerSpec' -notepropertyvalue ($sddcManagerObject | Select-Object -Skip 0)
        $managementDomainObject | Add-Member -notepropertyname 'networkSpecs' -notepropertyvalue $networkObject
        $managementDomainObject | Add-Member -notepropertyname 'nsxtSpec' -notepropertyvalue ($nsxtObject | Select-Object -Skip 0)
        $managementDomainObject | Add-Member -notepropertyname 'vsanSpec' -notepropertyvalue ($vsanObject | Select-Object -Skip 0)
        $managementDomainObject | Add-Member -notepropertyname 'dvsSpecs' -notepropertyvalue $dvsObject
        $managementDomainObject | Add-Member -notepropertyname 'clusterSpec' -notepropertyvalue ($clusterObject | Select-Object -Skip 0)
        $managementDomainObject | Add-Member -notepropertyname 'pscSpecs' -notepropertyvalue $pscObject
        $managementDomainObject | Add-Member -notepropertyname 'vcenterSpec' -notepropertyvalue ($vcenterObject | Select-Object -Skip 0)
        $managementDomainObject | Add-Member -notepropertyname 'hostSpecs' -notepropertyvalue $hostObject
        If (([INT]$commonObject.binarySettings.bringup.$($sharedRegionObject.release).cloudBuilderBuild) -ge "20355545") {
            If ($sharedRegionObject.subscriptionLicensing -eq "Included") {
                $managementDomainObject | Add-Member -notepropertyname 'subscriptionLicensing' -notepropertyvalue "True"
            }
            else {
                $managementDomainObject | Add-Member -notepropertyname 'subscriptionLicensing' -notepropertyvalue "False"
            }
        }

        Write-Host "Exporting the $module to $($path)$($pnpWorkbook.Workbook.Names["mgmt_sddc_domain"].Value)-domainSpec.json"
        $managementDomainObject | ConvertTo-Json -Depth 12 | Out-File -Encoding UTF8 -FilePath $path"$($pnpWorkbook.Workbook.Names["mgmt_sddc_domain"].Value)-domainSpec.json"
        Write-Host "Closing the Excel Workbook: $workbook"
        Close-ExcelPackage $pnpWorkbook -NoSave -ErrorAction SilentlyContinue
        Write-Host "Completed the Process of Generating the $module"
    }
    Catch {
        catchWriter -object $_
    }
}

Initialize-VCFManagementDomainSpecFromPnP -workbook ../../02-regiona-pnpWorkbook.xlsx -path ../../