# ETL RF Matrix Controller
Portable software to control ETL RF Matrix routers.

## Usage

When you open the application for the first time you will be greeted with an initial setup wizard. Here you can put the IP address of your router as well as the number of input and outputs. When you click "Continue" the program will generate the routable matrix.

From here you can click on the input labels on the left hand side to rename the labels and buttons. Across the top you can click on outputs to merge them into groups, rename and recolour the columns for easy differentiation.

## Features

- Customisable input and output ranges
- Groupable and colourable outputs
- Resizable window
- Highlight rows to differentiate inputs
- Customisable fonts for labels and buttons
- Customizable active route colours
- Toggle matrix crosshair
- Route multiple points at a time
- Save preset routes
- View matrix telemetry

## Installation from sources

Download the releases zip and extract to a folder on your computer.

### Windows
Open a command prompt in the extracted folder 

`cd path\to\files`

You may need to install PyInstaller if not already installed. 

`python pip install PyInstaller`

Run the following command 

`python -m PyInstaller --onefile --windowed --name "ETL Controller" --icon icon_1024.ico --add-data "icon_1024.ico;." etl_vortex_controller.py`

After the command has run you should see an "ETL Controller.exe" in the /dist folder.

### Mac OSX
Open terminal and cd to the extraced folder 

`cd path/to/files`

Make "build_macos.sh" executable with the following command 

`chmod 755 build_macos.sh`

Then run the following to build the application 

`./build_macos.sh`

One complete you should see an "ETL Controller" application in the /dist folder.

You can copy this file to your Applications folder so it is easily accessible.



