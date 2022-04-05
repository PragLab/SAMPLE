SAMPLE - Scanner Aquisition Manager Program for Laboratory Experiments

SAMPLE is a python script was designed to take time-lapse pictures from flat-bed scanners.

Usage and Installation:
	
	SAMPLE is available in two forms - as a standalone executable or the source python code.
	
	The standalone executable has no prerequisites in order to run. 
	
	To install the standalone executable, download and extract APP/SAMPLE.zip. 
	Then run SAMPLE.exe to open up the GUI and setup the time-lapse scan.
	'SAMPLE.ico' icon file must be placed in the same folder as the executable, for SAMPLE to work.
	
	To run the source code (SOURCE/SAMPLE.py), you must have a python interpeter version 3.6 or newer.
	You will also need to install Pillow if it is not installed in your python build.
	This can be done from the python console using the following command:
	
	pip install Pillow
	
	'SAMPLE.ico' is not required for the source code to run.
	
SAMPLE was written with python 3.8 and is compatible with windows 7, 8, 10 and 11.
Standalone executable was compiled using PyInstaller 4.10.

When executed, SAMPLE generates a GUI which allows the user to modify the intervals between each scan, 
the duration of the entire process and the format of the generated images.
Once initiated, a second window will open to monitor the progress of the time-lapse scan.

To work with SAMPLE, a scanner must have a WIA 2.0 compatible driver installed on the system,
otherwise the script will not be able to identify and connect to the scanner.

Image Formats produced by SAMPLE:

    BMP - Uncompressed bitmap, largest file size
    TIF - Lossless compression, large file size
    PNG - Lossless compression, small file size
    JPG - Lossy compression, smallest file size