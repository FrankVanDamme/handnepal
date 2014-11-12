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

find chocolib -name *.nupkg | head -n 1 | while read pkg 
do 
    file $pkg
    dir=$(dirname $pkg)
    basename=$(basename $pkg)
    mkdir work
    rm      ${basename}.zip
    cp $pkg ${basename}.zip
    cd work
    unzip ../${basename}.zip

    find -name '*.ps1' | while read ps1file
    do
	fromdos "$ps1file"
	grep -o "'http.*'" "$ps1file"   >> ../urls
	sed -i -e 's/http.*\/\([^\/]*$\)/s:\\chocolat\\packages\\\1/' "$ps1file" 
	todos "$ps1file"
    done
    read ee
    rm ../${basename}.zip
    zip -r ../${basename} *
    # ....
    cd ..
    # rm -R work 
done
