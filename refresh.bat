@ECHO OFF
C:
CD\
CLS


ECHO       .---.        .-----------
ECHO      /     \  __  /    ------
ECHO     / /     \(  )/    -----
ECHO    //////   ' \/ `   ---  RELEASE
ECHO   //// / // :    : ---    RENEW
ECHO  // /   /  /`    '--      FLUSHDNS     
ECHO //          //..\\        REGISTERDNS                   
ECHO        ====UU====UU====   DUMPS A LOG          
ECHO            '//II\\`       GPUPDATE FORCE
ECHO              ''``         RESETS PC
ECHO.
ECHO ======PRESS ANY KEY TO CONTINUE======
PAUSE > NUL

ipconfig /release 
timeout /t 1

ipconfig /renew 
timeout /t 1

ipconfig /flushdns 
timeout /t 1

ipconfig /registerdns 
timeout /t 2

netsh dump 
nbtstat -R 
netsh int ip reset C:\temp\reset.log 
netsh winsock reset 
timeout /t 1
shutdown /r /f /t 20
gpupdate /force /boot
timeout /t 1

