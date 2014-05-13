@echo off

SET CLASSPATH=lib\poi-3.10-FINAL-20140208.jar


SET CLASSPATH=%CLASSPATH%;.;

@rem echo %CLASSPATH%

javac -classpath %CLASSPATH% ConevitorUtil.java
javac -classpath %CLASSPATH% ConevitorMain.java
java -classpath %CLASSPATH% ConevitorMain
  
goto eof


:eof
ENDLOCAL