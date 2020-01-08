#!/bin/bash

OOXMLFILE=${1}
# OOXMLDIR=`basename ${OOXMLFILE} .xlsx`
OOXMLDIR=`basename ${OOXMLFILE}`
OOXMLDIR="${OOXMLDIR}.unpacked"

echo "unpacking file ${OOXMLFILE} into directory ${OOXMLDIR} ..."

/bin/rm -rf ${OOXMLDIR}
/bin/mkdir -p ${OOXMLDIR}

/usr/bin/unzip -o ${OOXMLFILE} -d ${OOXMLDIR}

for f in `find ./${OOXMLDIR} -name "*.xml"` ; do
  echo "formatting file $f ..."
  /usr/bin/tidy -quiet -xml -indent -wrap 60 -modify $f
done

for f in `find ./${OOXMLDIR} -name "*.rels"` ; do
  echo "formatting file $f ..."
  /usr/bin/tidy -quiet -xml -indent -wrap 60 -modify $f
done

