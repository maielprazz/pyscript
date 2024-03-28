# !/bin/bash
py $arg1 $arg2
py_status=$?
current_date=`date "+%D %T"`
msg="A script has ran to error
Server: $HOSTNAME
Script Name : $arg1
Time : $current_date"

echo "$py_status"
echo "$msg"

# msg="The following Script ran to error: $arg1 at $current_date"
# if [ $py_status -ne 0 ]; then
#     echo "$msg"
# fi;
