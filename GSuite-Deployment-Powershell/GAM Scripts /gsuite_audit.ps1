<#
.SYNOPSIS
This script is to further help automate the auditing of the clients G Suite environment

.DESCRIPTION
This script is to further help automate the auditing of the clients G Suite environment
There are a few requirements to get this script fully opperational
1. Minimum GAMADV-XTD3 version: 5.08.08
2. Ensure you have the correct information within your $HOME/.gam/gam.cfg file with the ability to <gam select client>

.PARAMETER cn
clientname as defined in your gam.cfg [Accepts multiple values.]

.PARAMETER sid
Optional. The Spreadsheet ID of the environment review sheet. Use only if running an environment audit for One client.

.PARAMETER test
Run a limited environment scan to test/debug the current script.

.INPUTS
None. You cannot pipe objects to gsuite_audit.ps1.

.OUTPUTS
System.String[]. Link(s) to the created Environment Review Spreadsheets

.EXAMPLE
PS> gsuite_audit.ps1 client_name
=====================================================
Results
=====================================================
client_name: Environment review link is https://drive.google.com/open?id=1tUu8E-4XSE838qdfdsmsjh9I5qHwKcht1qVCJzDvYWqIyjQ
==================================================================================

.EXAMPLE
PS> gsuite_audit.ps1 client_name -sid 1tUu8E-4XSE838qdfdsmsjh9I5qHwKcht1qVCJzDvYWqIyjQ
=====================================================
Results
=====================================================
client_name: Environment review link is https://drive.google.com/open?id=1tUu8E-4XSE838qdfdsmsjh9I5qHwKcht1qVCJzDvYWqIyjQ
==================================================================================

.EXAMPLE
PS> gsuite_audit.ps1 -cn client_name
=====================================================
Results
=====================================================
client_name: Environment review link is https://drive.google.com/open?id=1tUefsdfsu8E-4XSE838qmsjh9I5qHwKcht1qVCJzDvYWqIyjQ
==================================================================================

.EXAMPLE
PS> gsuite_audit.ps1 -cn client_name -sid 1tUefsdfsu8E-4XSE838qmsjh9I5qHwKcht1qVCJzDvYWqIyjQ
=====================================================
Results
=====================================================
client_name: Environment review link is https://drive.google.com/open?id=1tUefsdfsu8E-4XSE838qmsjh9I5qHwKcht1qVCJzDvYWqIyjQ
==================================================================================

.EXAMPLE
PS> gsuite_audit.ps1 -cn client_name,client_name2
=====================================================
Results
=====================================================
client_name: Environment review link is https://drive.google.com/open?id=1sersdfstUu8E-4XSE838qmsjh9I5qHwKcht1qVCJzDvYWqIyjQ
client_name2: Environment review link is https://drive.google.com/open?id=1SGysfewcfewfehzuP5vCc2SNNH7ZPxKr7o3zlpZhNGwAdejYUR_xk
==================================================================================

.EXAMPLE
PS> gsuite_audit.ps1 -cn client_name client_name2 client_name3
=====================================================
Results
=====================================================
client_name: Environment review link is https://drive.google.com/open?id=1tUu8Esfrewfesdf-4XSE838qmsjh9I5qHwKcht1qVCJzDvYWqIyjQ
client_name2: Environment review link is https://drive.google.com/open?id=1SGysfewsdsfsdfshzuP5vCc2SNNH7ZPxKr7o3zlpZhNGwAdejYUR_xk
client_name3: Environment review link is https://drive.google.com/open?id=hzuP5vCc2SNNH7ZPxKr7o3zlpZhNGwAdejYUR_xssdfkg
==================================================================================
#>


param (
        [Parameter(
			Mandatory=$true,
			Position = 0
		)]
		[string]$cn,

	    [Parameter(
			Mandatory=$false,
			Position = 1
		)]
		[string]$sid,

	    [Parameter(
            ValueFromRemainingArguments=$true,
            Position = 2
        )]
        [string[]]$listArgs,

		[Parameter(
			Mandatory=$false
		)]
		[Switch]$test
)




function New-Environment-Review-Sheet([string]$cn) {
	$sn = "[$cn] Environment Review"
	$timestampSn = "[$cn] Environment Review ($(Get-Date))"
	write-host "Creating a copy of the Environment Review template named $sn"
	# $sid = gam select $cn oauthuser copy drivefile 1doCnhUAtMtB-1RikiT1a4aazGlk-cILTdKiYCe_Acvc newfilename $timestampSn parentid root returnidonly
	$sid = gam select $cn oauthuser copy drivefile 1-nWoG3nYUqKO7aSOyxA-f87GNuabCDw8i3-_Lf9tXvw newfilename $timestampSn parentid root returnidonly
	return $sid
}

function Start-Environment-Audit-Test([string]$cn, [string]$sid) {
	$Watch = [System.Diagnostics.Stopwatch]::StartNew()
	$sn = "[$cn] Environment Review"

	write-host "Granting the cloudbakers.com domain Editor access to the Environment Review"
	gam select $cn redirect stdout null oauthuser add drivefileacl $sid domain cloudbakers.com role editor
	write-host "Collecting Org info and saving to $sn"
	gam select $cn info domain formatjson | ConvertFrom-Json | ConvertTo-Csv -NoTypeInformation | gam select $cn redirect stdout null oauthuser update drivefile $sid localfile - retainname gsheet "Org Info"
	Write-Host "====================================================="
	$Watch.Stop()
	Write-Host "[$cn] Execution took $($Watch.Elapsed.TotalSeconds) seconds"
	Write-Host "====================================================="
}

function Start-Environment-Audit([string]$cn, [string]$sid) {
	$Watch = [System.Diagnostics.Stopwatch]::StartNew()
	$sn = "[$cn] Environment Review"

	write-host "Granting the cloudbakers.com domain Editor access to the Environment Review"
	gam select $cn redirect stdout null oauthuser add drivefileacl $sid domain cloudbakers.com role editor

    try{
		write-host "Collecting Org info and saving to $($cn)_org_info.csv within the local directory"
		gam select $cn info domain formatjson | ConvertFrom-Json | Export-Csv ("$($cn)_org_info.csv")
		write-host "Collecting Org info and saving to $sn"
		gam select $cn info domain formatjson | ConvertFrom-Json | ConvertTo-Csv -NoTypeInformation | gam select $cn redirect stdout null oauthuser update drivefile $sid localfile - retainname gsheet "Org Info"
		#
		write-host "`r`nCollecting $cn's Domains and saving to $sn"
		gam select $cn print domains todrive tdsheet Domains tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true
		#
		write-host "`r`nCollecting $cn's Users and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet Users tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true print users licenses emailparts fields primaryemail,name,suspended,archived,ou,creationTime,lastLoginTime,gal
		#
		write-host "`r`nCollecting $cn's Addresses and saving to $sn"
		gam select $cn print addresses todrive tdsheet "All Addresses" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true
		#
		write-host "`r`nCollecting $cn's OUs and saving to $sn"
		gam select $cn print ous allfields todrive tdsheet OUs tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true
		#
		write-host "`r`nCollecting $cn's Groups and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet Groups tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true print groups allfields
		#
		write-host "`r`nCollecting $cn's Group Members and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Group Members" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true print group-members
		#
		write-host "`r`nCollecting $cn's Aliases and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet Aliases tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true print aliases
		#
		write-host "`r`nCollecting $cn's Delegates and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet Delegates tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print delegates
		#
		write-host "`r`nCollecting $cn's Send-as addresses and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Send-as addresses" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print sendas
		#
		write-host "`r`nCollecting $cn's Forwarding addresses and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Forwarding Addresses" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print forwardingaddresses
		#
		write-host "`r`nCollecting $cn's Forwarding Actions and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Forwards (Actions)" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print forwards enabledonly
		#
		write-host "`r`nCollecting $cn's POP email settings and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "POP" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print pop
		#
		write-host "`r`nCollecting $cn's IMAP email settings and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "IMAP" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print imap
		#
		write-host "`r`nCollecting $cn's Filters and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet Filters tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users print filters
		#
		write-host "`r`nCollecting $cn's Vacation (Out of Office) user settings and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Vacation (Out of Office)" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print vacation
		#
		write-host "`r`nCollecting $cn's Custom Admin Roles and saving to $sn"
		gam select $cn config csv_output_row_filter "'isSystemRole:boolean:false'" print roles privileges todrive tdsheet "Custom Admin Roles" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true
		#
		write-host "`r`nCollecting $cn's Admins and saving to $sn"
		gam select $cn print admins todrive tdsheet Admins tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true
		#
		write-host "`r`nCollecting $cn's Chrome Devices and saving to $sn"
		gam select $cn print cros allfields todrive tdsheet "Chrome Devices" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true
		#
		write-host "`r`nCollecting $cn's Resources and saving to $sn"
		gam select $cn print resources allfields todrive tdsheet Resources tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true
		#
		write-host "`r`nCollecting $cn's Buildings and saving to $sn"
		gam select $cn print buildings allfields todrive tdsheet Buildings tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true
		#
		write-host "`r`nCollecting $cn's Classic Sites and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Classic Sites" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print sites includeallsites withmappings roles all
		#
		write-host "`r`nCollecting $cn's New Sites and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "New Sites" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print filelist showmimetype gsite allfields
		#
		write-host "`r`nCollecting $cn's OAuth Tokens and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "OAuth Tokens" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true print tokens
		#
		write-host "`r`nCollecting $cn's Shared Drives - Domain and saving to $sn"
		gam select $cn print teamdrives todrive tdsheet "Shared Drives - Domain" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true
		#
		write-host "`r`nCollecting $cn's Shared Drives - User Access and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Shared Drives - User Access" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print teamdrives
		#
		write-host "`r`nCollecting $cn's Domain Shared Drives ACLs and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Domain Shared Drives ACLs" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true print teamdriveacls oneitemperrow
		#
		write-host "`r`nCollecting $cn's Secondary Calendars and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Secondary Calendars" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print calendars minaccessrole owner noprimary noresources nosystem nousers
		#
		write-host "`r`nCollecting $cn's Gmail and Drive storage used and saving to $sn"
		gam select $cn report users fields accounts:drive_used_quota_in_mb,accounts:gmail_used_quota_in_mb,accounts:total_quota_in_mb,accounts:used_quota_in_mb,accounts:used_quota_in_percentage todrive tdsheet "Storage Used" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true
		#
		write-host "`r`nCollecting $cn's Keep Notes and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Keep Notes" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print notes		
		#
		write-host "`r`nCollecting $cn's File Count and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "File Count" tdfileid $sid tdupdatesheet tdtitle $sn tdnobrowser true all users_ns_susp print filecounts showsize
		#
		write-host "`r`nCollecting $cn's Email Count and saving to $sn"
		gam select $cn config auto_batch_min 1 redirect csv - multiprocess todrive tdsheet "Email Count" tdfileid $sid tdupdatesheet tdtitle $sn all users_ns_susp print gmailprofile


		Write-Host "====================================================="
		$Watch.Stop()
		Write-Host "[$cn] Execution took $($Watch.Elapsed.TotalSeconds) seconds"
		Write-Host "====================================================="
	}
	catch{
		write-host "Something went wrong in single execution. $error"
		$error.clear()
	}
}

if(($null -eq $cn)){
	Write-host "You have not passed Client Name (-cn)"
}

else{
	$Watch = [System.Diagnostics.Stopwatch]::StartNew()
	$url_mapping = @{}
	try{
		$combined_args = $listArgs += $cn
		Write-Host "=================================================================================="
		if ($test) {
			Write-Host "Auditing '$($combined_args.Count)' environment(s) [TESTING]"
		} else {
			Write-Host "Auditing '$($combined_args.Count)' environment(s)"
		}
		Write-Host "=================================================================================="

		foreach ($client in $combined_args)
		{
			if (($sid) -and ($combined_args.Count -eq 1)) {
				$url_mapping[$client] = $sid
			} else {
				$url_mapping[$client] = New-Environment-Review-Sheet $client
			}


			if ($null -eq $url_mapping[$client]) {
				Write-Host "====================================================="
				Write-Error "Unable to create the Environment Review Sheet for [$client]. Skipping the review."
				$error.clear()
				continue
			}

			if ($test) {
				Start-Environment-Audit-Test $client $url_mapping[$client]
			} else {
				Start-Environment-Audit $client $url_mapping[$client]
			}
		}

		Write-Host "====================================================="
		Write-Host "Results"
		Write-Host "====================================================="
		$url_mapping.GetEnumerator() | ForEach-Object{
			if ($_.value -eq $null) {
				$message = '[{0}]: Unable to copy the environment review sheet.' -f $_.key
				Write-Host $message -ForegroundColor red
			} else {
				$message = '[{0}]: Environment review link is https://drive.google.com/open?id={1}' -f $_.key, $_.value
				Write-Host $message
			}
		}

		Write-Host "=================================================================================="
		$Watch.Stop()
		Write-Host "Program Execution took $($Watch.Elapsed.TotalSeconds) seconds"
		Write-Host "=================================================================================="
	}
	catch{
		write-host "Something went wrong in program execution. $error"
		$error.clear()
	}
}
