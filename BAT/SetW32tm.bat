#Echo off

w32tm /unregister

Timeout/T 5

net stop w32time

w32tm /register

Timeout /T 5

net start w32time

w32tm /config /syncfromflags:manual /manualpeerlist:time-a.nist.gov,time-b.nist.gov,time-c.nist.gov,time-d.nist.gov /update /reliable:yes