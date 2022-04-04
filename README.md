SAMPLE - Scanner Aquisition Manager Program for Laboratory Experiments

SAMPLE is a python script was designed to take time-lapse pictures from flat-bed scanners.

SAMPLE was written with python 3 and is compatible with windows 7, 8, 10 and 11.
Standalone executable was compiled using PyInstaller 4.10.

When executed, SAMPLE generates a GUI which allows the user to modify the intervals between each scan, 
the duration of the entire process and the format of the generated images.
Once initiated, a second window will open to monitor the progress of the time-lapse scan.

To work with SAMPLE, a scanner must have a WIA 2.0 compatible driver installed on the system,
otherwise the script will not be able to identify and connect to the scanner.

The script is available in two forms - as a standalone executable or the source python code.
'SAMPLE.ico' icon file is should be placed in the same folder as the script, for either version to work.

The standalone executable has no prerequisites in order to run. 
To run the source code, you must have a python interpeter version 3.6 or newer.

Image Formats produced by SAMPLE:

    BMP - Uncompressed bitmap, largest file size
    TIF - Lossless compression, large file size
    PNG - Lossless compression, small file size
    JPG - Lossy compression, smallest file size