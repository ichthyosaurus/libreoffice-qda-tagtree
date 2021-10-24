#!/bin/bash

echo "packaging..."
rm qda.oxt; zip -r qda.zip * -x build.sh _config.yml config.ini qda.oxt; mv qda.zip qda.oxt

echo "removing old version..."
unopkg remove qda.oxt

echo "installing new version..."
unopkg add qda.oxt

lowriter testdoc.odt
