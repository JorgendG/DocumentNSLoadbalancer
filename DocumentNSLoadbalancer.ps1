#Import-Module "$PSScriptRoot\adm.psm1"
#https://github.com/kbcitrite/scripts/tree/master/Citrix/ADM

if ( $null -eq $cred ) {
    $cred = Get-Credential -Message "adm01 credentials" -UserName nsroot
}

function GetRes ($ResType, $Name, $adchost) {
    if ( $null -eq $Name ) {
        (Invoke-ADMNitro -ADMSession $ADMSession -OperationMethod GET -ResourceType ("$ResType") -ADCHost $adchost).$ResType
    }
    else {
        (Invoke-ADMNitro -ADMSession $ADMSession -OperationMethod GET -ResourceType ("$ResType/$Name") -ADCHost $adchost).$ResType
    }
}

function Get-CertSubject ($subject) {
    $myregex = ',(?=(?:[^"]|"[^"]*")*$)'
    $parts = $subject -split $myregex

    $hash = $null
    $hash = @{}

    foreach ($part in $parts) {
        $splitpart = $part -split "="
        if ( ($splitpart[0]).trim() -ne "DC" ) {
            $hash.add( ($splitpart[0]).trim() , $splitpart[1] )
        }
    }

    return $hash
}


$ADMSession = Connect-ADM -ADMHost https://adm01.homelabdc22.local -Cred $cred #-ApiVersion v1

$ns_lbvservers = (Invoke-ADMNitro -ADMSession $ADMSession -OperationMethod GET -ResourceType 'ns_lbvserver').ns_lbvserver
"======"
#$ns_lbvservers = $ns_lbvservers | where{ $_.name -like '*mediawiki*' }
$ns_lbvservers = $ns_lbvservers | Where-Object { $_.name -like '*web*' }
$ns_lbvservers | Select-Object name, vsvr_type, state, ns_ip_address
$ns_lbvservers
"======"
""
New-WordInstance -WindowState wdWindowStateNormal
New-WordDocument

$FirstPage = $true

foreach ($ns_lbvserver in $ns_lbvservers) {
    if ( -not $FirstPage ) {
        Add-WordBreak NewPage
        
    }
    $FirstPage = $false
    $lbvserver = GetRes -ResType "lbvserver" -Name "$($ns_lbvserver.name)" -ADCHost $ns_lbvserver.ns_ip_address

    Add-WordText "$($lbvserver.name) - $($ns_lbvserver.ns_ip_address)" -WDBuiltinStyle wdStyleHeading1
    
    $lbprops = @(
        [pscustomobject]@{Col1 = 'Name'; Col2 = $lbvserver.name; Col3 = 'Listen Priority'; Col4 = $lbvserver.priority }
        [pscustomobject]@{Col1 = 'Protocol'; Col2 = $lbvserver.servicetype; Col3 = 'Listen Policy Expression'; Col4 = $lbvserver.listenpolicy }
        [pscustomobject]@{Col1 = 'State'; Col2 = $lbvserver.curstate; Col3 = 'Redirection Mode'; Col4 = $lbvserver.m }
        [pscustomobject]@{Col1 = 'IP Address'; Col2 = $lbvserver.ipv46; Col3 = 'Range'; Col4 = $lbvserver.range }
        [pscustomobject]@{Col1 = 'Port'; Col2 = $lbvserver.port; Col3 = 'IPset'; Col4 = 'XXXX' }
        [pscustomobject]@{Col1 = 'Trafic Domain'; Col2 = $lbvserver.td; Col3 = 'RHI State'; Col4 = $lbvserver.rhistate }
        [pscustomobject]@{Col1 = 'Comment'; Col2 = $lbvserver.comment; Col3 = 'AppFlow Logging'; Col4 = $lbvserver.appflowlog }
        [pscustomobject]@{Col1 = ''; Col2 = ''; Col3 = 'Retain Connection on Cluster'; Col4 = $lbvserver.retainconnectionsoncluster }
        [pscustomobject]@{Col1 = ''; Col2 = ''; Col3 = 'Redirect From Port'; Col4 = $lbvserver.redirectfromport }
        [pscustomobject]@{Col1 = ''; Col2 = ''; Col3 = 'HTTPS Redirect URL'; Col4 = $lbvserver.httpsredirecturl }
        [pscustomobject]@{Col1 = ''; Col2 = ''; Col3 = 'TCP Probe Port'; Col4 = 'XXXX' }

    )

    $lbprops

    Add-WordTable -Object $lbprops -GridTable 'Grid Table 1 Light' -HeaderRow:$false -FirstColumn:$false -RemoveProperties

    $lbvserver_binding = GetRes -ResType lbvserver_binding -Name "$($ns_lbvserver.name)" -ADCHost $ns_lbvserver.ns_ip_address
    $bindings = $lbvserver_binding | Get-Member -Force | Where-Object { $_.membertype -eq 'NoteProperty' } | Select-Object name

    foreach ( $binding in $bindings.name ) {
        switch ($binding) {
            'lbvserver_service_binding' {
                ""
                "-- Service Binding --"
                $lbvserver_service_binding = $lbvserver_binding.lbvserver_service_binding | Select-Object servicename, ipv46, port, servicetype, curstate
                $lbvserver_service_binding
                Add-WordTable -Object $lbvserver_service_binding -GridTable 'Grid Table 1 Light'-FirstColumn:$false
                
            }
            'lbvserver_servicegroup_binding' {
                ""
                "-- Servicegroup Binding --"
                $lbvserver_binding.lbvserver_servicegroup_binding | Select-Object servicegroupname | Format-Table
                $sgs = $lbvserver_binding.lbvserver_servicegroup_binding
                foreach ( $sg in $sgs ) {
                    $sg_lbmonitor_binding = GetRes -ResType "servicegroup_lbmonitor_binding" -Name "$($sg.servicegroupname)?view=detail" -ADCHost $ns_lbvserver.ns_ip_address
                    $sg_lbmonitor_binding
                    if ( $sg_lbmonitor_binding ) {
                        Add-WordTable -Object $sg_lbmonitor_binding -GridTable 'Grid Table 1 Light'-FirstColumn:$false
                    }
                }

            }
            'lbvserver_servicegroupmember_binding' {
                ""
                "-- Servicegroup Member Binding --"
                $lbvserver_binding.lbvserver_servicegroupmember_binding | Select-Object servicegroupname, ipv46, port, servicetype, curstate | Format-Table
                
                $sgmembers = $lbvserver_binding.lbvserver_servicegroupmember_binding | Select-Object @{ N = 'Servicegroup'; e = { $_.servicegroupname } },
                @{ N = 'IP Address'; e = { $_.ipv46 } },
                @{ N = 'Port'; e = { $_.port } },
                @{ N = 'Servicetype'; e = { $_.servicetype } },
                @{ N = 'Current State'; e = { $_.curstate } }
                Add-WordTable -Object $sgmembers -GridTable 'Grid Table 1 Light'-FirstColumn:$false
            }
            'lbvserver_rewritepolicy_binding' {
                ""
                "-- Rewrite Policy --"
                $lbvserver_binding.lbvserver_rewritepolicy_binding | Format-Table
                $rwpols = $lbvserver_binding.lbvserver_rewritepolicy_binding | Select-Object policyname, priority, gotopriorityexpression
                

                $rwpols = $lbvserver_binding.lbvserver_rewritepolicy_binding | Select-Object @{ N = 'Policy'; e = { $_.policyname } },
                @{ N = 'Priority'; e = { $_.priority } },
                @{ N = 'GoTo'; e = { $_.gotopriorityexpression } }
                $rwpols
                Add-WordTable -Object $rwpols -GridTable 'Grid Table 1 Light'-FirstColumn:$false
            }
            'lbvserver_auditsyslogpolicy_binding' {
                ""
                "-- Auditsyslog Policy --"
                $lbvserver_binding.lbvserver_auditsyslogpolicy_binding | Select-Object policyname | Format-Table
            }
            'lbvserver_csvserver_binding' {
                ""
                "-- Content Switching Policy --"
                $lbvserver_csvserver_binding = $lbvserver_binding.lbvserver_csvserver_binding | Select-Object cachevserver, priority, hits
                $lbvserver_csvserver_binding | Format-Table

                Add-WordTable -Object $lbvserver_csvserver_binding -GridTable 'Grid Table 1 Light'-FirstColumn:$false
            }
            'Name' {}

                       
            Default {
                ""
                Write-Host "Unknown Binding $binding" -ForegroundColor Red
            }
        }

    }


     
    #region SSL Server
    if ( $lbvserver.servicetype -eq 'SSL' ) {
        "-- Certificate --"
        $sslvserver_sslcertkey_binding = GetRes -ResType "sslvserver_sslcertkey_binding" -Name "$($ns_lbvserver.name)" -ADCHost $ns_lbvserver.ns_ip_address
        $sslcertkey = GetRes -ResType "sslcertkey" -Name "$($sslvserver_sslcertkey_binding[0].certkeyname)" -ADCHost $ns_lbvserver.ns_ip_address
        $sslcertkey = $sslcertkey | Select-Object @{ N = 'Certificate'; e = { $sslcertkey.certkey } },
        @{ N = 'Cert file'; e = { $sslcertkey.cert } },
        @{ N = 'Key file'; e = { $sslcertkey.key } },
        @{ N = 'Subject'; e = { (Get-CertSubject $sslcertkey.subject).CN } },
        @{ N = 'Issued By'; e = { (Get-CertSubject $sslcertkey.servername).CN } },
        @{ N = 'Days valid'; e = { $sslcertkey.daystoexpiration } } 
        $sslcertkey
        Add-WordTable -Object $sslcertkey -GridTable 'Grid Table 1 Light'-FirstColumn:$true -HeaderRow:$false -VerticleTable

        "-- SSL Profile --"
        $sslvserver = GetRes -ResType sslvserver -Name "$($ns_lbvserver.name)?view=detail" -ADCHost $ns_lbvserver.ns_ip_address
        $sslvserver = $sslvserver | Select-Object  @{ N = 'SSL Profile'; e = { $_.sslprofile } }
        $sslvserver
        Add-WordTable -Object $sslvserver -GridTable 'Grid Table 1 Light'-FirstColumn:$false -HeaderRow:$true

        $sslvserver_sslciphersuite_bindings = GetRes -ResType sslvserver_sslciphersuite_binding -Name "$($ns_lbvserver.name)" -ADCHost $ns_lbvserver.ns_ip_address
        $sslvserver_sslciphersuite_bindings = $sslvserver_sslciphersuite_bindings | Select-Object  @{ N = 'Cipher Name'; e = { $_.ciphername } },
        @{ N = 'Description'; e = { $_.description } }
        $sslvserver_sslciphersuite_bindings
        Add-WordTable -Object $sslvserver_sslciphersuite_bindings -GridTable 'Grid Table 1 Light'-FirstColumn:$false -HeaderRow:$true
    }
    #endregion SSL Server

    "-- Loadbalance Method --"
    $lbmethod = GetRes -ResType lbvserver -Name "$($ns_lbvserver.name)?view=detail" -ADCHost $ns_lbvserver.ns_ip_address
    $lbmethod = $lbmethod | Select-Object  @{ N = 'Loadbalance Method'; e = { $_.lbmethod } }
    $lbmethod
    Add-WordTable -Object $lbmethod -GridTable 'Grid Table 1 Light'-FirstColumn:$false -HeaderRow:$true

    "-- Session Persistence --"
    $persistencetype = GetRes -ResType lbvserver -Name "$($ns_lbvserver.name)?view=detail" -ADCHost $ns_lbvserver.ns_ip_address
    $persistencetype = $persistencetype | Select-Object  @{ N = 'Session Persistence'; e = { $_.persistencetype } }
    $persistencetype
    Add-WordTable -Object $persistencetype -GridTable 'Grid Table 1 Light'-FirstColumn:$false -HeaderRow:$true

    $dnsrecords = Resolve-DnsName $lbvserver.ipv46 -Server '192.168.1.22'
    $dnsrecords = $dnsrecords | Select-Object @{ N = 'IP Address'; e = { "$($lbvserver.ipv46)" } }, @{ N = 'FQDN'; e = { "$($_.NameHost)" } }
    $dnsrecords
    Add-WordTable -Object $dnsrecords -GridTable 'Grid Table 1 Light'-FirstColumn:$false
    
    #GetRes -ResType lbparameter -ADCHost $ns_lbvserver.ns_ip_address


    <#
    Close-WordDocument -SaveOptions wdDoNotSaveChanges
    Close-WordInstance
    #>
    
}
   