"""
Created on March 3 2022

Author: Elon Yariv

Version: 1.12

SAMPLE - Scanner Aquisition Manager Program for Laboratory Experiments

SAMPLE is a python script was designed to take time-lapse pictures from flat-bed scanners.

Usage and Installation:
	
	SAMPLE is available in two forms - as a standalone executable or the source python code.
	
	The standalone executable has no prerequisites in order to run. 
	
	To install the standalone executable, download and extract APP/SAMPLE.zip. 
	Then run SAMPLE.exe to open up the GUI and setup the time-lapse scan.
	'SAMPLE.ico' icon file must be placed in the same folder as the executable, for SAMPLE to work.
	
    Before running SAMPLE, the scanner must have a WIA compatible driver installed on the system, 
    otherwise the scanner will not be recognized by the script.
    
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

Image Formats produced by SAMPLE:

    BMP - Uncompressed bitmap, largest file size
    TIF - Lossless compression, large file size
    PNG - Lossless compression, small file size
    JPG - Lossy compression, smallest file size
    
Version Histroy:
    Version 1.12:
    - Added 'refresh' button for the scanner list.
    - Several visual rearangements of the windows.
    Version 1.11:
    - Added 'view in folder' button for the scanning monitor screen.
    Version 1.1:
    - Reworked the scheduler logic, it is more accurate and will no longer skip the occational  scan.
    - Fixed an issue with the attributes of currently operating runs, changing when user changes something in the main window.
    
"""

import time, os, re, threading, pythoncom
import win32com.client as win32
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
from PIL import Image,ImageTk

# Global Variable to track if the computer is scanning

IsScanning = False

class ImageScanner:
    def __init__(self, parent, args):
        self.parent = parent
        child = self.child = tk.Toplevel(parent)
        self.input_table = args.copy()
        
        child.title("SAMPLE - " + self.input_table["Scanner"])
        child.geometry("665x365")
        child.resizable(0,0)
        child.protocol("WM_DELETE_WINDOW", lambda:self.ConformationWindow(2))
        
        self.inum = 0
        self.IsNotPaused = True
        self.hasFinished = False
        
        if self.input_table["Delay"] > 0:
            self.onDelay = True
        else:
            self.onDelay = False
        
        self.start_time = time.time()
        self.scan_time = round(self.start_time + int(self.input_table["Delay"])*60)
        self.pause_time = 0
        
        border = tk.Canvas(child,width = 680,height = 360)
        border.create_rectangle(5,230,420,360,outline = 'grey', width=1)
        border.create_rectangle(435,26,658,360,outline = 'grey', width=1)
        border.place(x = 0, y = 0)
        
        self.log = scrolledtext.ScrolledText(child, wrap = tk.WORD, height=11, width=50)
        self.log.bind("<Key>", lambda e: "break")
        self.log.place(x = 5, y = 30)
        
        # Preview Image
        
        self.canvas_size = [210, 297]
        self.canvas = tk.Canvas(child ,width = self.canvas_size[0], height = self.canvas_size[1], bg = 'white')
        self.canvas.create_text(self.canvas_size[0]/2, self.canvas_size[1]/2, text="Preview", fill="#e6e6e6", font=('Helvetica 32 bold'))
        self.canvas.place(x = 440, y = 30)
        
        tk.Label(child, text="Scanning Log:").place(x = 5, y = 5)
        tk.Label(child, text="Last Scanned Image:").place(x = 430, y = 5)
        
        self.Progress2Scan = ttk.Progressbar(child, orient = tk.HORIZONTAL, 
                                           length = 300, mode = 'determinate')
        self.P2Slabel = tk.Label(child, text = self.update_progress(True))
        
        self.Progress2Scan.place(x = 10, y = 240)
        self.P2Slabel.place(x = 310 , y = 240)
        
        self.progress = ttk.Progressbar(child, orient = tk.HORIZONTAL, 
                                           length = 300, mode = 'determinate')
        self.TPlabel = tk.Label(child, text = self.update_progress(False))
        
        self.progress.place(x = 10, y = 290)
        self.TPlabel.place(x = 310, y = 290)
        
        self.bStop = tk.Button(child, text = "Pause", command = lambda:self.ConformationWindow(0))
        self.bContinue = tk.Button(child, text = "Continue", command = lambda:self.ConformationWindow(1))
        self.bExit = tk.Button(child, text = "Exit", command = lambda:self.ConformationWindow(2))
        self.bShowFolder = tk.Button(child, text = "View In Folder", command = lambda:self.openfile(self.input_table["Output"]))
        
        self.bStop.place(x = 10, y = 330, width = 60)
        self.bContinue.place(x = 120, y = 330, width = 60)
        self.bExit.place(x = 340, y = 330, width = 60)
        self.bShowFolder.place(x = 510, y = 330, width = 80)
        
        self.bContinue["state"] = "disabled"
        
        self.input_table["log"] = self.input_table["Output"] + '/' + self.input_table["Name"] + ".log"
        
        with open(self.input_table["log"], 'w') as logfile:
            current_date = time.strftime("%m/%d/%y - %H:%M - ",time.localtime())
            lines = [current_date + "Scanning has been initiated.\n",
                     "Output will be written to - " + self.input_table["Output"] + "\n",
                     "Resolution: " + str(round(self.input_table["Width"],1)) + "x" + str(round(self.input_table["Height"],1)) + " inch\n",
                     "DPI: " + str(self.input_table["dpi"]) + "   Color: " + self.input_table["Colour"] + "   Format: " + self.input_table["Format"] + "\n",
                     "\n"]
            for line in lines:
                self.log.insert(tk.END, line)
                logfile.write(line)
        
        child.after(0, self.scantimer)
        
    def scantimer(self):
        if round(time.time()) >= self.scan_time:
            if self.onDelay:
                self.onDelay = False
            
            self.inum += 1 # Integer storing how many scans were performed
            
            thread = threading.Thread(target = self.InitScan, args = [self.input_table])
            thread.start() # Initiate Scan
            
            # Reset progress till next scan
            # Set time when next scan will be performed
            
            self.Progress2Scan["value"] = 0
            self.scan_time = round(self.start_time + self.inum*self.input_table["Interval"]*60 + self.input_table["Delay"]*60 + self.pause_time)
            
        if self.IsNotPaused:
            if self.inum < self.input_table["Repetitions"]:
                self.P2Slabel["text"] = self.update_progress(True)
                self.TPlabel["text"] = self.update_progress(False)
                self.child.after(1000, self.scantimer)
            else: # time-lapse scan is finished
                with open(self.input_table["log"], 'w') as logfile:
                    current_date = time.strftime("%m/%d/%y - %H:%M - ",time.localtime())
                    line = current_date + "Finished scanning all %s images.\n" %self.inum
                    
                    self.log.insert(tk.END, line)
                    logfile.write(line)
                    
                self.TPlabel["text"] = self.update_progress(False)
                    
                self.inum = 0
                self.bStop["state"] = "disabled"
                self.bContinue["state"] = "disabled"
                self.hasFinished = True
                self.IsNotPaused = False
    
    def update_progress(self, istime2scan = True):
        time_diff = round(time.time()-self.start_time-self.pause_time)
        
        if istime2scan:
            if self.onDelay:
                interval = self.input_table["Delay"]*60
                time2scan = (time_diff)%(int(interval))
            else:
                interval = self.input_table["Interval"]*60
                time2scan = (time_diff-self.input_table["Delay"]*60)%(int(interval))
            
            self.Progress2Scan["value"] = (time2scan/interval)*100
            
            return f"Next Scan - {int((interval-time2scan)/60)}:{int((interval-time2scan)%60):02d}"
        else:
            total_time = self.input_table["Interval"]*(self.input_table["Repetitions"]-1)*60+self.input_table["Delay"]*60
            time_left = total_time - time_diff
            
            self.progress["value"] = 100 - (time_left/total_time)*100
            
            return f"Time left - {int(time_left/3600)}:{int((time_left%3600)/60):02d}:{int(time_left%60):02d}"
    
    def InitScan(self, args):
        global IsScanning
        
        while IsScanning:
            time.sleep(0.5) # Sleep while another scanner is active
              
        try:
            IsScanning = True
            self.bExit["state"] = "disabled"
            
            self.canvas.delete("all")
            self.canvas.create_text(self.canvas_size[0]/2, self.canvas_size[1]/2, text="Scanning in\n  Progress", fill="green", font=('Helvetica 24 bold'))
            
            pythoncom.CoInitialize()
        
            dm = win32.Dispatch("WIA.DeviceManager")
            ip = win32.Dispatch("WIA.ImageProcess")
            
            WIA_COMMAND_TAKE_PICTURE = "{AF933CAC-ACAD-11D2-A093-00C04F72DC3C}"
            
            if args["Colour"] == "RGB":
                colour_code = 1
                colour_depth = 24
            elif args["Colour"] == "Greyscale":
                colour_code = 2
                colour_depth = 8
            elif args["Colour"] == "Black&White":
                colour_code = 4
                colour_depth = 1
            else:
                colour_code = 1
                colour_depth = 24
            
            if args["Format"] == "TIFF":
                imgFormat = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
            elif args["Format"] == "BMP":
                imgFormat = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
            elif args["Format"] == "PNG":
                imgFormat = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
            elif args["Format"] == "JPG":
                imgFormat = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
            else:
                imgFormat = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
        
            # Go over each connected WIA device.
            # If it has the same name as the selected scanner, connect to this device.
            
            for info in dm.DeviceInfos:
                for prop in info.Properties:
                    if prop.Name == "Name" and prop.Value == args["Scanner"]:
                        Scanner = info.Connect()
            
            if Scanner:
                PropDict = dict()
                
                PropDict["Horizontal Resolution"] = args["dpi"]
                PropDict["Vertical Resolution"] = args["dpi"]
                PropDict["Current Intent"] = colour_code
                PropDict["Bits Per Pixel"] = colour_depth
                PropDict["Horizontal Extent"] = args["Width"]*args["dpi"] # Width in pixels
                PropDict["Vertical Extent"] = args["Height"]*args["dpi"] # Height in pixels
                
                # for debugging purposes - lists all the properties of the scanner
                #
                #for i, p in enumerate(Scanner.Items[Scanner.Items.count].Properties):
                #    print(i, p.Name, p.PropertyID, p.Value)
                
                for prop in Scanner.Items[Scanner.Items.count].Properties:
                    for key in PropDict:
                        if key == prop.Name:
                            prop.Value = PropDict[key]
                
                for command in Scanner.Commands:
                    if command.CommandID == WIA_COMMAND_TAKE_PICTURE:
                        Scanner.ExecuteCommand(WIA_COMMAND_TAKE_PICTURE)
            
                # No matter what format is given here, WIA will produce an uncompressed BMP file.
                tmp_img = Scanner.Items[Scanner.Items.Count].Transfer(imgFormat)
            
                # Convert the original BMP file into the desired format.
                ip.Filters.add(ip.FilterInfos("Convert").FilterID)
                ip.Filters[Scanner.Items.count].Properties["FormatID"] = imgFormat
                image = ip.apply(tmp_img)
                
                filename = self.get_filename(args["Name"], args["Format"])
                fullpath = args["Output"] + '/' + filename
                tmppath = args["Output"] + '/' + "temp.bmp"
            
                if os.path.exists(fullpath):
                    os.remove(fullpath)
            
                # Save scanned image to output folder.
                
                image.SaveFile(fullpath)
            
                # Tkinter crashes when attempting to display the TIF files created by WIA 2.0.
                # I suspect this might be due to the lossless LZW compression.
                # If TIF format was chosen, save a temporary copy, using the original BMP file.
            
                if args["Format"] == 'TIFF':
                    if os.path.exists(tmppath):
                        os.remove(tmppath)
                    
                    tmp_img.SaveFile(tmppath)
                
                ScanFinished = True
            else:
                ScanFinished = False
        except:
            ScanFinished = False
                
        IsScanning = False
        self.bExit["state"] = "normal"
        self.canvas.delete("all")
        
        if ScanFinished:
            try:
                if args["Format"] == 'TIFF':
                    IMG = Image.open(tmppath)
                else:
                    IMG = Image.open(fullpath)
                    
                IMG = IMG.resize((self.canvas_size[0],self.canvas_size[1]), Image.ANTIALIAS)
                self.preview = ImageTk.PhotoImage(image = IMG)
                self.canvas.create_image((0,0),anchor = tk.NW, image = self.preview)
                line = "Image written to " + filename + "\n"
                if args["Format"] == 'TIFF':
                    os.remove(tmppath) # After preview is uploaded to the GUI, delete temporary file.
            except FileNotFoundError:
                self.canvas.create_text(self.canvas_size[0]/2, self.canvas_size[1]/2, text="File not\n Found", fill="red", font=('Helvetica 32 bold'))
                line = "Output file" + filename + " not found.\n"
            
            with open(self.input_table["log"], 'a') as logfile:
                self.log.insert(tk.END, line)
                logfile.write(line)
        else:
            self.canvas.create_text(self.canvas_size[0]/2, self.canvas_size[1]/2, text="Scanner\n Error", fill="red", font=('Helvetica 32 bold'))
    
    def DoStuff (self, args):
        pass
        
    def get_filename(self, filename, filetype):
        current_date = time.strftime("-%m%d%y-%H%M",time.localtime())
        return filename + current_date + '.' + filetype.lower()
    
    def ExitWindow(self, wintype = 0):
        if wintype == 0:
            with open(self.input_table["log"], 'a') as logfile:
                current_date = time.strftime("%m/%d/%y - %H:%M - ",time.localtime())
                line = current_date + "Terminated by user - %s images scanned.\n" %self.inum
                logfile.write(line)
            self.child.destroy()
        elif wintype == 1:
            self.grandchild.destroy()
    
    def stop(self):
        self.pause_start = time.time()
        
        self.IsNotPaused = False
        self.toggle_buttons()
        with open(self.input_table["log"], 'a') as logfile:
            current_date = time.strftime("%m/%d/%y - %H:%M - ",time.localtime())
            line = current_date + "Scanning has been halted.\n"
            self.log.insert(tk.END, line)
            logfile.write(line)
        self.ExitWindow(1)
    
    def resume(self):
        self.pause_time += time.time() - self.pause_start
        
        self.IsNotPaused = True
        self.toggle_buttons()
        with open(self.input_table["log"], 'a') as logfile:
            current_date = time.strftime("%m/%d/%y - %H:%M - ",time.localtime())
            line = current_date + "Scanning has resumed.\n"
            self.log.insert(tk.END, line)
            logfile.write(line)
            
        self.child.after(1000, self.scantimer)
        self.ExitWindow(1)
    
    def openfile(self, directory):
        try:
            os.startfile(directory)
        except FileNotFoundError:
            print (directory, " not found.")
    
    def toggle_buttons(self):
        buttons = [self.bStop,self.bContinue]
        for b in buttons:
            if b["state"] == "normal":
                b["state"] = "disabled"
            elif b["state"] == "disabled":
                b["state"] = "normal"
                
    def ConformationWindow(self, wintype = 0):
        self.grandchild = tk.Toplevel(self.child)
        self.grandchild.geometry("300x100")
        self.grandchild.resizable(0,0)
        self.grandchild.protocol("WM_DELETE_WINDOW", self.disable_event)
        
        if wintype == 0:
            self.grandchild.title("Stop Scan")
            tk.Label(self.grandchild, text = "Do you really want to pause the scan on\n" + self.input_table["Scanner"] + "?").place(relx = 0.15, rely = 0.1)
            tk.Button(self.grandchild, text = 'Yes', command = self.stop).place(relx = 0.05, rely = 0.6, relwidth = 0.2)
            tk.Button(self.grandchild, text = 'No', command = lambda:self.ExitWindow(1)).place(relx = 0.75, rely = 0.6, relwidth = 0.2)
        elif wintype == 1:
            self.grandchild.title("Resume Scan")
            tk.Label(self.grandchild, text = "Resume scanning with\n" + self.input_table["Scanner"] + "?").place(relx = 0.3, rely = 0.1)
            tk.Button(self.grandchild, text = 'Yes', command = self.resume).place(relx = 0.05, rely = 0.6, relwidth = 0.2)
            tk.Button(self.grandchild, text = 'No', command = lambda:self.ExitWindow(1)).place(relx = 0.75, rely = 0.6, relwidth = 0.2)
        elif wintype == 2:
            self.grandchild.title("Terminate Scan")
            tk.Label(self.grandchild, text = "Do you really want to exit?").place(relx = 0.25, rely = 0.1)
            tk.Button(self.grandchild, text = 'Yes', command = lambda:self.ExitWindow(0)).place(relx = 0.05, rely = 0.6, relwidth = 0.2)
            tk.Button(self.grandchild, text = 'No', command = lambda:self.ExitWindow(1)).place(relx = 0.75, rely = 0.6, relwidth = 0.2)
        
    def disable_event(self):
        pass
        
class SAMPLE:
    def __init__(self, master):
        
        self.master = master
        
        self.Argtable = dict()
        self.Argtable["Output"] = None
        self.MaxSize = [None, None]
        
        master.geometry("510x380")
        
        self.Argtable["Scanner"] = None
        self.update_scanner_list()
        
        self.outputdir = None
        
        master.protocol("WM_DELETE_WINDOW", lambda:self.ConformationWindow(1))
        master.title("SAMPLE - Scanner Aquisition Manager Program for Lab Experiments")
        master.resizable(0, 0)
        
        # Create grey background rectangles 
        
        border = tk.Canvas(master,width = 510,height = 380)
        border.create_rectangle(5,22,225,225,outline = 'grey', width=1) # Scanner list
        border.create_rectangle(256,52,500,145,outline = 'grey', width=1) # Scan options
        border.create_rectangle(256,185,500,225,outline = 'grey', width=1) # Image format options
        border.create_rectangle(5,250,500,340,outline = 'grey', width=1) # File name & folder
        border.create_line(10,130,220,130,fill = 'grey', width=1) # Line in scanner list box
        border.place(x = 0, y = 0)
        
        tk.Label(master, text="Version 1.12", fg = 'grey').place(x = 440, y = 0)
        
        self.dirbox = tk.Text(master, height=1, width=40)
        self.dirbox.bind("<Key>", lambda e: "break")
        self.dirbox.place(x = 80, y = 306)

        self.bStart = tk.Button(master, text = "Start Scan", command = lambda:self.ConformationWindow(0))
        self.bStart.place(x = 10, y = 348, width = 80)
        
        self.bExit = tk.Button(master, text = "Exit Program", command = lambda:self.ConformationWindow(1))
        self.bExit.place(x = 420, y = 348, width = 80)
        
        self.bExplore = tk.Button(master, text = "Browse", command = self.BrowseFiles)
        self.bExplore.place(x = 410, y = 303, width = 80)
        
        tk.Label(master, text="Directory:").place(x = 15, y = 305)
        tk.Label(master, text="Output Options").place(x = 10, y = 228)
        
        vcmd = master.register(self.validate)
        nvcmd = master.register(self.validate_name)
        fvcmd = master.register(self.validate_float)
        
        tk.Label(master, text="Scan Settings").place(x = 260, y = 30)
        
        self.nInput = tk.Entry(master, validate = 'key', validatecommand = (vcmd, '%P'))
        self.nInput.insert(tk.END, '49')
        self.nInput.place(x = 340, y = 60, width = 50)
        tk.Label(master, text="Repetitions:").place(x = 260, y = 60)
        
        self.iInput = tk.Entry(master, validate = 'key', validatecommand = (vcmd, '%P'))
        self.iInput.insert(tk.END, '60')
        self.iInput.place(x = 340, y = 90, width = 50)
        tk.Label(master, text="Interval:").place(x = 260, y = 90)
        tk.Label(master, text="Minutes").place(x = 390, y = 90)
        
        self.delayInput = tk.Entry(master, validate = 'key', validatecommand = (vcmd, '%P'))
        self.delayInput.insert(tk.END, '0')
        self.delayInput.place(x = 340, y = 120, width = 50)
        tk.Label(master, text="Start After:").place(x = 260, y = 120)
        tk.Label(master, text="Minutes").place(x = 390, y = 120)
        
        self.widthInput = tk.Entry(master, validate = 'key', validatecommand = (fvcmd, '%P'))
        self.widthInput.place(x = 60, y = 140, width = 100)
        tk.Label(master, text="Width:").place(x = 15, y = 140)
        tk.Label(master, text="Inches").place(x = 160, y = 140)
        
        self.heightInput = tk.Entry(master, validate = 'key', validatecommand = (fvcmd, '%P'))
        self.heightInput.place(x = 60, y = 170, width = 100)
        tk.Label(master, text="Height:").place(x = 15, y = 170)
        tk.Label(master, text="Inches").place(x = 160, y = 170)
        
        self.resInput = tk.Entry(master, validate = 'key', validatecommand = (vcmd, '%P'))
        self.resInput.insert(tk.END, '300')
        self.resInput.place(x = 60, y = 200, width = 100)
        tk.Label(master, text="Pixels").place(x = 160, y = 200)
        tk.Label(master, text="DPI:").place(x = 15, y = 200)
        
        self.nameInput = tk.Entry(master, validate = 'key', validatecommand = (nvcmd, '%P'))
        self.nameInput.place(x = 80, y = 260, width = 324)
        tk.Label(master, text="File Name:").place(x = 15, y = 258)
        
        # Scanner selection menu
        
        self.ScanBox = tk.Listbox(master, listvariable=self.ScannerList, height = 4, selectmode = 'browse', exportselection=False) 
        scrollbar = ttk.Scrollbar(master, orient='vertical', command=self.ScanBox.yview)
        self.ScanBox.config(yscrollcommand = scrollbar.set)
        self.ScanBox.bind('<<ListboxSelect>>', self.select_item)
        
        self.ScanBox.place(x = 10, y = 28, width = 195, height = 70)
        scrollbar.place(x = 205, y = 28, height = 70)
        tk.Label(master, text="Available Scanners").place(x = 10, y = 1)
        
        tk.Button(master, text = "Refresh", command = self.update_scanner_list).place(x = 160, y = 100, width = 60)
        
        # Image format configuration
        
        tk.Label(master, text="Image Format Options").place(x = 260, y = 163)
        
        colours = ["RGB", "Greyscale", "Black&White"]
        
        self.selected_colour = tk.StringVar(master,colours[0])
        self.ColourMenu = tk.OptionMenu(master, self.selected_colour, *colours)
        self.ColourMenu.configure(width = '10')
        self.ColourMenu.place(x = 260, y = 190)
        
        formats = ["TIFF","BMP","PNG","JPG"]
        
        self.selected_format = tk.StringVar(master,formats[0])
        self.FormatMenu = tk.OptionMenu(master, self.selected_format, *formats)
        self.FormatMenu.configure(width = '10')
        self.FormatMenu.place(x = 390, y = 190)
        
        self.bStart["state"] = "disabled"
        self.widthInput["state"] = "disabled"
        self.heightInput["state"] = "disabled"
        
        self.check_can_initiate()
    
    def update_scanner_list(self):
        self.ScannerList = tk.StringVar(value=self.get_available_scanners())
    
    def get_available_scanners(self):
        Scannerlist = []
        dm = win32.Dispatch("WIA.DeviceManager")
        for info in dm.DeviceInfos:
            #if info.Type == 1:
            for prop in info.Properties:
                if prop.Name == "Name":
                    Scannerlist.append(prop.Value)
        return Scannerlist
                    
    
    def BrowseFiles(self):
        self.Argtable["Output"] = filedialog.askdirectory(initialdir = os.path.expanduser('~'), title = "Select folder for output")
        if self.Argtable["Output"]:
            self.dirbox.delete(1.0,tk.END)
            self.dirbox.insert(tk.END, self.Argtable["Output"])
    
    def ConformationWindow(self, wintype = 0, err = None):
        self.SubWindow = tk.Toplevel(self.master)
        self.SubWindow.geometry("300x100")
        self.SubWindow.resizable(0,0)
        self.SubWindow.protocol("WM_DELETE_WINDOW", self.disable_event)
        
        if wintype == 0:
            self.SubWindow.title("Start Scan")
            tk.Label(self.SubWindow, text = "Do you really want to start scanning on\n" + self.Argtable["Scanner"] + "?").place(relx = 0.15, rely = 0.1)
            tk.Button(self.SubWindow, text = 'Yes', command = self.start).place(relx = 0.05, rely = 0.6, relwidth = 0.2)
            tk.Button(self.SubWindow, text = 'No', command = lambda:self.ExitWindow(1)).place(relx = 0.75, rely = 0.6, relwidth = 0.2)
        elif wintype == 1:
            self.SubWindow.title("Close Program")
            tk.Label(self.SubWindow, text = "Do you really want to exit?\nThis will terminate all active scans.").place(relx = 0.15, rely = 0.1)
            tk.Button(self.SubWindow, text = 'Yes', command = lambda:self.ExitWindow(0)).place(relx = 0.05, rely = 0.6, relwidth = 0.2)
            tk.Button(self.SubWindow, text = 'No', command = lambda:self.ExitWindow(1)).place(relx = 0.75, rely = 0.6, relwidth = 0.2)
        elif wintype == 2:
            self.SubWindow.title("Error")
            tk.Label(self.SubWindow, text = err).place(relx = 0.5, rely = 0.1, anchor = tk.N)
            tk.Button(self.SubWindow, text = 'OK', command = lambda:self.ExitWindow(1)).place(relx = 0.5, rely = 0.65, relwidth = 0.2, anchor = tk.N)
    
    def ExitWindow(self, wintype = 0):
        if wintype == 0:
            self.master.destroy()
        elif wintype == 1:
            self.SubWindow.destroy()
    
    def update_progressbar (self):
        return f"{round(self.progress['value'],1)}%"
    
    def start(self):
        self.SubWindow.destroy()
        try:
            # Register the arguments given by the user and start the scan.
            
            self.Argtable["Width"] = float(self.widthInput.get())
            self.Argtable["Height"] = float(self.heightInput.get())
            self.Argtable["dpi"] = int(self.resInput.get())
            self.Argtable["Delay"] = int(self.delayInput.get())
            self.Argtable["Interval"] = int(self.iInput.get())
            self.Argtable["Repetitions"] = int(self.nInput.get())
            self.Argtable["Name"] = str(self.nameInput.get())
            self.Argtable["Colour"] = str(self.selected_colour.get())
            self.Argtable["Format"] = str(self.selected_format.get())
            
            if self.Argtable["Width"] > self.MaxSize[0] or self.Argtable["Height"] > self.MaxSize[1]:
                self.ConformationWindow(2, f"Image size is larger then the selected scanner's tray.\nMaximal size is {self.MaxSize[0]}x{self.MaxSize[1]} inches.")
            elif self.Argtable["Width"] <= 0 or self.Argtable["Height"] <= 0:
                self.ConformationWindow(2, f"Image size should be a positive number.\nMaximal size is {self.MaxSize[0]}x{self.MaxSize[1]} inches.")
            else:
                ImageScanner(self.master, self.Argtable)
        except ValueError:
            self.ConformationWindow(2, f"Image size is not a valid number.\nMaximal size is {self.MaxSize[0]}x{self.MaxSize[1]} inches.")
        
    def select_item(self, event):
        self.ScanBox = event.widget
        if self.ScanBox.curselection() != ():
            self.Argtable["Scanner"] = str(self.ScanBox.get(self.ScanBox.curselection()))
        
            dm = win32.Dispatch("WIA.DeviceManager")
        
            for info in dm.DeviceInfos:
                for prop in info.Properties:
                    if prop.Name == "Name" and prop.Value == self.Argtable["Scanner"]:
                        Scanner = info.Connect()
        
            for prop in Scanner.Items[Scanner.Items.count].Properties:
                if prop.Name == "Horizontal Extent":
                    self.MaxSize[0] = prop.Value/100
                
                    self.widthInput["state"] = "normal"
                    self.widthInput.delete(0, tk.END)
                    self.widthInput.insert(tk.END, self.MaxSize[0])
                elif prop.Name == "Vertical Extent":
                    self.MaxSize[1] = prop.Value/100
                
                    self.heightInput["state"] = "normal"
                    self.heightInput.delete(0, tk.END)
                    self.heightInput.insert(tk.END, self.MaxSize[1])
        
    def validate(self, P):
        if str.isdigit(P) or P == "":
            return True
        else:
            return False
    
    def validate_float(self, P):
        pattern = re.compile(r"^[0-9]+\.?[0-9]*")
        if re.match(pattern, P) or P == '':
            return True
        else:
            return False
    
    def validate_name(self, P):
        if str.isalnum(P) or P == '':
            return True
        else:
            return False
    
    def check_can_initiate(self):
        if len(self.nameInput.get()) > 0 and self.Argtable["Output"] and self.Argtable["Scanner"] and int(self.iInput.get()) > 0 and int(self.nInput.get()):
            self.bStart["state"] = "normal"
        else:
            self.bStart["state"] = "disabled"
        self.master.after(1000, self.check_can_initiate)
            
    def disable_event(self):
        pass
        
root = tk.Tk()
app = SAMPLE(root)
root.mainloop()