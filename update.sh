for i in `find . -type f | grep xlsx$`
do
	name=`dirname $i`/`basename $i .xlsx`
	dest="$name.zip"
	cp -f $i $dest
	rm -rf $name
	unzip -q $dest -d $name
	rm -f $dest
done
