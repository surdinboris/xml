#!/bin/bash


function ServersScan {
	
	rm -fr $HomeDir/ActiveServers_${HostName}.log 2>/dev/null
	echo -en "\nScanning system for active servers, please wait"
	waitdots
	
	UpHosts=$(nmap -sP 20.0.0.1-14|awk '/appears to be up/ {print $2}')
	if [ "$(echo $UpHosts|wc -w)" -ne "0" ]; then
		echo_color bold "$(echo $UpHosts|wc -w) active server/s found"
		echo -en "Discovering their model type, product number and S/N, please wait"
		waitdots
	fi
	
	for IP in $UpHosts ;do 
		echo "$IP" >> $HomeDir/ActiveServers_${HostName}.log

	done
	
	if [ -s $HomeDir/ActiveServers_${HostName}.log ]; then
		ServersCount=$(cat $HomeDir/ActiveServers_${HostName}.log|wc -l)
		echo -e "\n$ServersCount servers are alive:"
		echo "---------------------"
		echo -e "   IP\t\tSystem Type\t  S/N\t     P/N"
		echo -e "   --\t\t-----------\t  ---\t     ---"
		cat $HomeDir/ActiveServers_${HostName}.log|column -s\- -t
	else
		echo -e "\033[31m\nNo active servers found, please check!\e[0m"
		echo -e "Do you want to rescan?"
		select Option in $(echo "Yes Exit"); do
			if [ "$Option" == "Yes" ]; then
				skip=""
				ServersScan
				break
			elif [ "$Option" == "Exit" ]; then
				dhcpservice stop
				echo "Exiting..."
				exit 10
			fi
		done
	fi
	
	echo -e "\nDo you want to continue or wait for more servers to respond?"
}


############ MAIN ############

unset SESSION_MANAGER # dismiss vnc related terminal messages
HostName=`hostname`
DateStr=$(date +"%Y%m%d_%H%M%S")
rm -fr $HomeDir/RunStatus_${HostName}.log 2>/dev/null
touch $HomeDir/RunStatus_${HostName}.log
OSRelease=$(lsb_release -a|awk '/Release/ {print $2}')
LogDir="$HomeDir/Logs"
Log="$LogDir/runnning-$DateStr.log"
touch $Log

Main > >(tee -i "$Log") 2>&1

