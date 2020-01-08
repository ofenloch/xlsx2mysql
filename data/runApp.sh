#!/bin/bash

JARFILE="./target/xls2mysql-1.1-SNAPSHOT.jar"
MAINCLASS="de.ofenloch.xlsx2mysql.App"
DEPENDENCIES="./target/dependency"

JAVA=/usr/lib/jvm/default-java/bin/java

JAVAARGS="-agentlib:jdwp=transport=dt_socket,server=n,suspend=y,address=localhost:38031 -Dfile.encoding=UTF-8"
JAVAARGS="-Dfile.encoding=UTF-8"

CLASSPATH=${JARFILE}

for jar in ${DEPENDENCIES}/*.jar ; do
    if [ -f "${jar}" ] ; then
        #echo "adding ${jar} to CLASSPATH"
        CLASSPATH=${CLASSPATH}:${jar}
    fi
done

# echo "CLASSPATH is ${CLASSPATH}"


echo "calling \"${JAVA} ${JAVAARGS} -classpath ${CLASSPATH} ${MAINCLASS} $@\" ..."
${JAVA} ${JAVAARGS} -classpath ${CLASSPATH} ${MAINCLASS} $@