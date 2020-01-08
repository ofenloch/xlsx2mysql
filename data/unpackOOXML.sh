#!/bin/bash

DOCXFILE=${1}
# DOCXDIR=`basename ${DOCXFILE} .xlsx`
DOCXDIR=`basename ${DOCXFILE}`
DOCXDIR="${DOCXDIR}.unpacked"

echo "unpacking file ${DOCXFILE} into directory ${DOCXDIR} ..."

/bin/rm -rf ${DOCXDIR}
/bin/mkdir -p ${DOCXDIR}

/usr/bin/unzip -o ${DOCXFILE} -d ${DOCXDIR}

for f in `find ./${DOCXDIR} -name "*.xml"` ; do
  echo "formatting file $f ..."
  /usr/bin/tidy -quiet -xml -indent -wrap 60 -modify $f
done

for f in `find ./${DOCXDIR} -name "*.rels"` ; do
  echo "formatting file $f ..."
  /usr/bin/tidy -quiet -xml -indent -wrap 60 -modify $f
done

