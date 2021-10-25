#!/bin/bash

echo "packaging..."
rm qda.oxt
zip -r qda.zip * -x build.sh _config.yml config.ini qda.oxt \
    TODO testdoc.* \
    src/pythonpath/hsluv/{MANIFEST.in,README.md,setup.cfg,setup.py} \
    src/pythonpath/hsluv/tests/* src/pythonpath/hsluv/tests \
    '.*' '*/.*' '*~' '*/*~'
mv qda.zip qda.oxt

echo "removing old version..."
unopkg remove qda.oxt

echo "installing new version..."
unopkg add qda.oxt

lowriter testdoc.odt
