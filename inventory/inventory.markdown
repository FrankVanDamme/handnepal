# While running a Linux live cdrom/usb drive

Capture output of these commands (use "script" to capture terminal session): 
* cat /proc/cpuinfo
* free
* sudo dmidecode
* lspci
* lsusb
* sudo lshw
* inxi -Fx -c 0    
    * (in dropbox)



# While running on a presumably installed Windows OS

* run ocs inventory agent, and import the files into OCS. Need to remove afterwards (and need ocs server to import).

Importing is done with /usr/share/ocsinventory-server/binutils/ocsinventory-injector.pl --directory
