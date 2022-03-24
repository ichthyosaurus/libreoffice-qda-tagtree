#!/bin/bash
# This file is part of libreoffice-qda-tagtree.
# SPDX-FileCopyrightText: 2021-2022 Mirian Margiani
# SPDX-License-Identifier: GPL-3.0-or-later

if [[ "$1" == "-h" || "$1" == "--help" ]]; then
    echo "usage: build.sh [install|install-test|i|it]"
fi

cBUILD_BASE="build"
cBUILD="$cBUILD_BASE/qda-tagtree"

echo "packaging..."

if [[ -e "$cBUILD_BASE" ]]; then
    if ! mv "$cBUILD_BASE" "$cBUILD_BASE.old" --backup=t; then
        echo "error: failed to remove old build directory"
        exit 1
    fi
fi

mkdir -p "$cBUILD_BASE"

cp -r qda-tagtree "$cBUILD"
cp README.md "$cBUILD"
rm -rf "$cBUILD"/src/pythonpath/hsluv/{MANIFEST.in,README.md,setup.cfg,setup.py,.git,.gitignore,.travis.yml}
rm -rf "$cBUILD"/src/pythonpath/hsluv/tests

find "$cBUILD" -iname "*~" -delete
find "$cBUILD" -iname ".*" -delete

back="$(pwd)"
cd "$cBUILD"
zip -r qda-tagtree.zip *

cd "$back"
mv "$cBUILD/qda-tagtree.zip" "$cBUILD_BASE/qda-tagtree.oxt"


if [[ "$1" == "install"* || "$1" == "i"* ]]; then
    echo "removing old version..."
    unopkg remove qda.oxt

    echo "installing new version..."
    unopkg add qda.oxt
fi

if [[ "$1" == "install-test" || "$1" == "it" ]]; then
    lowriter testdoc.odt
fi
