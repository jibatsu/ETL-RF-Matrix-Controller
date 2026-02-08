#!/bin/bash

# Build the app
pyinstaller --onefile --windowed \
    --name "ETL Controller" \
    --icon icon.icns \
    --add-data "icon.icns:." \
    etl_vortex_controller.py

# Code sign the app (to avoid Gatekeeper warnings)
codesign --force --deep --sign - "dist/ETL Controller.app"