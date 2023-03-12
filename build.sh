#!/bin/bash

rm -rf ./build/
mkdir ./build/

cp -r ./assets/. ./build/

cd cmd/cli

go build -o i2e .

cd ../..

cp cmd/cli/i2e ./build/