@echo off

SET CLASSPATH=lib\poi-3.10-FINAL-20140208.jar


SET CLASSPATH=%CLASSPATH%;.;

echo %CLASSPATH%

javac -classpath %CLASSPATH% Conevitor.java
java -classpath %CLASSPATH% Conevitor
  
goto eof


:eof
ENDLOCAL