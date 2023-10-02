#Requires -modules CredentialManager
#Import-Module "$PSScriptRoot\adm.psm1"

#https://github.com/kbcitrite/scripts/tree/master/Citrix/ADM


$cred = Get-StoredCredential -Target admnsroot

if ( $null -eq $cred ) {
    $cred = Get-Credential -UserName "nsroot" -Message "ADM Credentials"
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

$ns_vpnvservers = (Invoke-ADMNitro -ADMSession $ADMSession -OperationMethod GET -ResourceType 'ns_vpnvserver').ns_vpnvserver
"======"
$ns_vpnservers | Select-Object name, vsvr_type, state, ns_ip_address
#$ns_lbvservers
"======"
""

foreach ($ns_vpnvserver in $ns_vpnvservers) {
    $vpnvserver = GetRes -ResType "vpnvserver" -Name "$($ns_vpnvserver.name)" -ADCHost $ns_vpnvserver.ns_ip_address

    $lbprops = @(
        [pscustomobject]@{Col1 = 'Name'; Col2 = $ns_vpnvserver.name; Col3 = 'Listen Priority'; Col4 = $ns_vpnvserver.priority }
        [pscustomobject]@{Col1 = 'Protocol'; Col2 = $ns_vpnvserver.vsvr_type; Col3 = 'Listen Policy Expression'; Col4 = $ns_vpnvserver.listenpolicy }
        [pscustomobject]@{Col1 = 'State'; Col2 = $ns_vpnvserver.state; Col3 = 'Redirection Mode'; Col4 = $ns_vpnvserver.m }
        [pscustomobject]@{Col1 = 'IP Address'; Col2 = $ns_vpnvserver.vsvr_ip_address; Col3 = 'Range'; Col4 = $ns_vpnvserver.range }
        [pscustomobject]@{Col1 = 'Port'; Col2 = $ns_vpnvserver.vsvr_port; Col3 = 'IPset'; Col4 = 'XXXX' }
        [pscustomobject]@{Col1 = 'Trafic Domain'; Col2 = $ns_vpnvserver.td; Col3 = 'RHI State'; Col4 = $ns_vpnvserver.rhistate }
        [pscustomobject]@{Col1 = 'Comment'; Col2 = $ns_vpnvserver.comment; Col3 = 'AppFlow Logging'; Col4 = $ns_vpnvserver.appflowlog }
        [pscustomobject]@{Col1 = ''; Col2 = ''; Col3 = 'Retain Connection on Cluster'; Col4 = $ns_vpnvserver.retainconnectionsoncluster }
        [pscustomobject]@{Col1 = ''; Col2 = ''; Col3 = 'Redirect From Port'; Col4 = $ns_vpnvserver.redirectfromport }
        [pscustomobject]@{Col1 = ''; Col2 = ''; Col3 = 'HTTPS Redirect URL'; Col4 = $ns_vpnvserver.httpsredirecturl }
        [pscustomobject]@{Col1 = ''; Col2 = ''; Col3 = 'TCP Probe Port'; Col4 = 'XXXX' }

    )
    $lbprops | ft


    $vpnvserver_binding = GetRes -ResType vpnvserver_binding -Name "$($ns_vpnvserver.name)" -ADCHost $ns_vpnvserver.ns_ip_address
    $bindings = $vpnvserver_binding | Get-Member -Force | Where-Object { $_.membertype -eq 'NoteProperty' } | Select-Object name

    foreach ( $binding in $bindings.name ) {
        switch ($binding) {
            'vpnvserver_vpnsessionpolicy_binding' {
                ""
                "-- Rewrite Policy --"
                $vpnvserver_binding.vpnvserver_vpnsessionpolicy_binding | Format-Table
                $rwpols = $vpnvserver_binding.lbvserver_rewritepolicy_binding | Select-Object policyname, priority, gotopriorityexpression
                

                $rwpols = $vpnvserver_binding.lbvserver_rewritepolicy_binding | Select-Object @{ N = 'Policy'; e = { $_.policyname } },
                @{ N = 'Priority'; e = { $_.priority } },
                @{ N = 'GoTo'; e = { $_.gotopriorityexpression } }
                $rwpols
                #Add-WordTable -Object $rwpols -GridTable 'Grid Table 1 Light'-FirstColumn:$false
            }
            'vpnvserver_staserver_binding' {
                ""
                "-- vpnvserver_staserver_binding Policy --"
                $vpnvserver_binding.vpnvserver_staserver_binding | Format-Table
            }
            'vpnvserver_cachepolicy_binding' {
                ""
                "-- vpnvserver_cachepolicy_binding Policy --"
                $vpnvserver_binding.vpnvserver_cachepolicy_binding | Format-Table
            }
            'Name' {}

                       
            Default {
                ""
                Write-Host "Unknown Binding $binding" -ForegroundColor Red
            }
        }

    }


     
    #region SSL Server
    if ( $vpnvserver.servicetype -eq 'SSL' ) {
        "-- Certificate --"
        $sslvserver_sslcertkey_binding = GetRes -ResType "sslvserver_sslcertkey_binding" -Name "$($vpnvserver.name)" -ADCHost $vpnvserver.ns_ip_address
        $sslcertkey = GetRes -ResType "sslcertkey" -Name "$($sslvserver_sslcertkey_binding[0].certkeyname)" -ADCHost $vpnvserver.ns_ip_address
        $sslcertkey = $sslcertkey | Select-Object @{ N = 'Certificate'; e = { $sslcertkey.certkey } },
        @{ N = 'Cert files'; e = { "$($sslcertkey.cert) $($sslcertkey.key)" } },
        #@{ N = 'Key file'; e = { $sslcertkey.key } },
        @{ N = 'Subject'; e = { (Get-CertSubject $sslcertkey.subject).CN } },
        @{ N = 'Issued By'; e = { (Get-CertSubject $sslcertkey.servername).CN } },
        @{ N = 'Days valid'; e = { $sslcertkey.daystoexpiration } } 
        $sslcertkey
        #Add-WordTable -Object $sslcertkey -GridTable 'Grid Table 1 Light'-FirstColumn:$true -HeaderRow:$false -VerticleTable

        "-- SSL Profile --"
        $sslvserver = GetRes -ResType sslvserver -Name "$($vpnvserver.name)?view=detail" -ADCHost $vpnvserver.ns_ip_address
        $sslvserver = $sslvserver | Select-Object  @{ N = 'SSL Profile'; e = { $_.sslprofile } }
        $sslvserver
        #Add-WordTable -Object $sslvserver -GridTable 'Grid Table 1 Light'-FirstColumn:$false -HeaderRow:$true

        $sslvserver_sslciphersuite_bindings = GetRes -ResType sslvserver_sslciphersuite_binding -Name "$($vpnvserver.name)" -ADCHost $vpnvserver.ns_ip_address
        $sslvserver_sslciphersuite_bindings = $sslvserver_sslciphersuite_bindings | Select-Object  @{ N = 'Cipher Name'; e = { $_.ciphername } },
        @{ N = 'Description'; e = { $_.description } }
        $sslvserver_sslciphersuite_bindings
        #Add-WordTable -Object $sslvserver_sslciphersuite_bindings -GridTable 'Grid Table 1 Light'-FirstColumn:$false -HeaderRow:$true
    }
    #endregion SSL Server

    
}
   