#!/bin/sh

# # # # # # # # # #
# XBRL_NI_GENERATOR
# # # # # # # # # #

CLASSPATH=./XBRLToolNI.jar;./lib/*

# echo ${CLASSPATH}

java -cp ${CLASSPATH} it.infocamere.xbrl.ni.ui.UIlauncher ./resources/log4j.properties 2018-11-04 1>log/log.out 2>log/log.err
