#!/bin/bash

rm -rf ./build/
mkdir ./build/

cp -r ./assets/. ./build/

echo Building CLI...

cd cmd/cli
go build -o i2e .
cd ../..

echo Building web..
cd cmd/web
go build -o i2e_web .
cd ../..

cp cmd/cli/i2e ./build/

cp -r cmd/web/public ./build/public
cp cmd/web/i2e_web ./build/