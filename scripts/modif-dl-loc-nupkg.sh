#!/bin/bash

for binary in fromdos todos 
do
    which $binary
    if [[ $? -ne 0 ]]
    then 
	echo $binary not found! install tofrodos
	exit 1
    fi
done

rm -R work 
rm urls

find /home/shared/chocolatey/lib -iname '*.nupkg' | while read pkg 
do 
    dir=$(dirname $pkg)
    basename=$(basename $pkg)
    echo ZIP file: $basename
#########continue
    mkdir work
    rm      ${basename}.zip
    echo copying...
    cp -v $pkg ${basename}.zip
    echo ... ok 
    cd work
    unzip -q ../${basename}.zip

    find -name '*.ps1' | while read ps1file
    do
	echo PS1 FILE $ps1file found
	fromdos "$ps1file"
	grep -o "'http.*'" "$ps1file"   >> ../urls
	sed -i -e 's/http.*\/\([^\/]*$\)/s:\\chocolat\\packages\\\1/' "$ps1file" 
	todos "$ps1file"
    done
    zip -qr ../${basename} *
    # ....
    cd ..
    mv urls "urls$(date +%F_%U-%M)"
    mv -v ${basename} $pkg 
    rm  ${basename}.zip 
    rm -R work 
done
rm nupkgs
