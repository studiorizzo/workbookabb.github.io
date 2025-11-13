@echo off

REM # # # # # # # # # # 
REM  XBRL_NI_GENERATOR
REM # # # # # # # # # # 

set CLASSPATH=./XBRLToolNI.jar;./lib/*

REM echo CLASSPATH=%CLASSPATH%

java -classpath %CLASSPATH% it.infocamere.xbrl.ni.ui.UIlauncher ./resources/log4j.properties 2018-11-04 1>log/log.out 2>log/log.err

