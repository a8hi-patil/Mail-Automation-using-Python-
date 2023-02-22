#! /bin/bash
x=$(date +%s) 
dif='7'
x=$((x - $dif * 24 * 60 * 60))
fdate=$(date --date @$x +%Y/%m/%d)
ftime="-00:00:00"
FromDate="$fdate"
echo $FromDate
#*****************Date difference**********
y=$(date +%s) 
dif2='1'
y=$((y - $dif2 * 24 * 60 * 60))
tdate=$(date --date @$y +%Y/%m/%d)
ttime="-23:59:59"
ToDate="$tdate"
echo $ToDate
Year=$(date --date @$x +%Y)
Month_name=$(date --date @$x +%b)
Date=$(date +"%d")
now="7 days ago"
first=$(date -d "$now - $(($(date +%_d -d "$now" )-1)) days")
WOM=$(( 1 + 10#$(date +%U -d "$now") - 10#$(date -d "$first" +%U) ))

#*********************************Directories****************************************
bin=/data/bin
data=/data/daily_data
data_dir=/data/database/$Year/$Month_name/$WOM

Conversion_LOG_Dir=/data/log
Log_File_Name="${Month_name}_W${WOM}"
echo `date +%d_%m_%y_%H_%M_%S` " *************************************">> $Conversion_LOG_Dir/$Log_File_Name.log


mkdir -p $data_dir
cp -r $data/*.csv $data_dir
cd $data
rm -f *.csv

cd $data_dir
mkdir Dir1 Dir2
find . -name '*dir1.csv' | cpio -pdm $data_dir/Dir1/
find . -name '*dir2.csv' | cpio -pdm $data_dir/Dir2/



for X in Dir1 Dir2 
do  	
	Dir=$data_dir/$X/
	if [ $X == "TYPE" ]
	then
		BU="BU"
	elif [ $X == "TYPE2" ]
	then
		BU="TYPE3"
	else
		BU="TYPE4"
	fi

	if [ "$(ls -A $Dir)" ]
	then 
		
		PDFName=$Month_name"_W"$WOM"_"$Year"_"$BU"_"$X
		Output_Path=$data_dir/$X/$PDFName.pdf
		python3.8 $bin/MakeReportPDF.py $Dir $Output_Path $X $BU $FromDate $ToDate
		cd $data_dir/$VT/
		cp /data/workArea/MIS_Report.xlsx .
		python3.8 $bin/DVA_MIS_MailSend.py $Dir $PDFName $VT $BU $WOM $Month_name $Year
		
	else
		echo `date +%d_%m_%y_%H_%M_%S` " *****-- $Dir is  Empty. Mail not sent for $BU $VT DVA Report">> $Conversion_LOG_Dir/$Log_File_Name.log	
	fi
done
