#
#	find-wd15.ps1
#	By:  van Gennip, Brett
#	Comments:  This script is created to track the
#				The Dell wd15 docking stations on the network
#				
#
#  I am not responsible for any damages.  Run at own risk.
#
#	Date: Apr 25, 2019
#
##############################


# Insert your text file with a list of hostnames on your network

$x=$(cat c:\temp\748090s.txt)

#
#   This function checks to see if the hostname is responding to ICMP PING
#
function ison1($xx){	
	test-connection -count 1  $xx -quiet
	return $?
}

#
#    Creates the CSV file to write to
#
if(-not $(test-path c:\temp\list.csv)){
Write-output "ComputerName,Date,Status,Model,Dock-Mac-Address" |Out-File c:\temp\list.csv
}
#for each of these hostnames in the text file do the following:

foreach($xx in $x){

$a=ison1($xx)
#Send 1 icmp ping test
	if(-not $a[0]){
			#hostname, time/date, host not reachable --> record in list
		
			Write-output "$xx,$(get-date -format d),offline,offline,offline" |Out-File -Append c:\temp\list.csv -noclobber
				}
	if($a[0]){
			
			$model=""
			$model=$(gwmi win32_computersystem -computername $xx).model
		#if wmi is not responding, mark as gwmi no connection	
	if($model){
			$docmac=""
			$docmac=$(gwmi -query "SELECT * FROM Win32_NetworkAdapter WHERE ServiceName = 'rtux64w10'" -computername $xx).macaddress
			
			if(-not $docmac){
				$docmac=$(gwmi -query "SELECT * FROM Win32_NetworkAdapter WHERE ServiceName = 'rtux64w7'" -computername $xx).macaddress
						if(-not $docmac){
							$docmac="No Dock connected"
						}
			}
			
			Write-output "$xx,$(get-date -format d),online,$model,$docmac" |Out-File -Append c:\temp\list.csv -noclobber
		}else{
		Write-output "$xx,wmi-fail,$(get-date -format d),online,wmi-fail,wmi-fail" |Out-File -Append c:\temp\list.csv -noclobber
		}
	}
}
