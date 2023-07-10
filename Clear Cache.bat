@ECHO OFF
C:
CD\
CLS


ECHO       .---.        .-----------
ECHO      /     \  __  /    ------
ECHO     / /     \(  )/    -----CLEAR IE + JAVA CACHE
ECHO    //////   ' \/ `   ---   
ECHO   //// / // :    : ---     
ECHO  // /   /  /`    '--       
ECHO //          //..\\         
ECHO        ====UU====UU====            
ECHO            '//II\\`        
ECHO              ''``           
ECHO.
ECHO ======PRESS ANY KEY TO CONTINUE======
PAUSE > NUL

cd c:\windows\system32
javaws -uninstall
ECHO Clearing Java Cache
timeout /t 1

RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8
ECHO Clearing Temporary Internet Files: 
timeout /t 1

RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1
ECHO Clearing History:  
timeout /t 1

RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2
ECHO Clearing Cookies:  
timeout /t 1

ECHO 
ECHO 
ECHO 
ECHO 
ECHO 
ECHO                  Complete
ECHO 
ECHO 
ECHO 
exit