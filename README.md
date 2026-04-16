# LabVIEW-BmC-DAQ
Contains custom built LabVIEW application for human data collection during bimanual coordination tasks. The following details installation, setup, and basic usage of the graphical user interface (GUI).

## Installation
This project was created using [LabVIEW 2020 (64-bit)](https://www.ni.com/en/support/downloads/software-products/download.labview.html#585712). Use this version for best results. Changing versions may break the code or cause other bugs.  

### Required Software and Dependencies
This project requires the following additional software and packages to run. 

* [VI Package Manager 2020](https://www.vipm.io/): Used to manage LabVIEW packages. The Free version was used for this project. Required packages are listed below and can be installed using VIPM:
  * [MGI library](https://www.vipm.io/package/mgi_lib_mgi_library/)
  * [MGI Actor Framework Message Maker](https://www.vipm.io/package/mgi_lib_mgi_actor_framework_message_maker/)
  * [MGI Monitored Actor](https://www.vipm.io/package/mgi_lib_monitored_actor/)
  * [OpenG Toolkit](https://www.vipm.io/package/openg.org_lib_openg_toolkit/)
  * [Viewpoint TSVN Toolkit](https://www.vipm.io/package/viewpoint_tsvn_toolkit/) (optional if using Tortoise SVN for source control)
  * [JKI_tool_tortoisesvn](https://www.vipm.io/package/jki_tool_tortoisesvn/) (optional if using Tortoise SVN for source control)
* [Phidgets drivers](https://www.phidgets.com/docs/Operating_System_Support): Required to use Phidgets accelerometer device.
* [Phidgets LabVIEW libraries](https://www.phidgets.com/docs/Language_-_LabVIEW): Required to develop in LabVIEW with Phidgets devices. Even if not using Phidgets devices, this library is required to run the code.
* [EMGworks](https://delsys.com/support/emgworks/): Required to acquire data from EMG sensor base.
* [NI-DAQ<sup>TM</sup>mx](https://www.ni.com/en/support/downloads/drivers/download.ni-daq-mx.html#585854): Drivers required for data acquisition in LabVIEW with NI devices. Note: Install programming environments such as NI LabVIEW or Microsoft Visual Studio before installing this product
* [Microsfot Excel](https://www.microsoft.com/en-us/microsoft-365/excel): Required to final data output
* [TDM Excel Add-In for Microsoft Excel](https://www.ni.com/en/support/downloads/tools-network/download.tdm-excel-add-in-for-microsoft-excel.html#378046): Required to convert temporary TDMS files to final output Excel file. This add-in may need to be enabled in Excel to work.
* [Tortise SVN](https://tortoisesvn.net/) (optional): Used for version control during development. If the user wants to make additional changes to the program, Tortoise SVN or another source control service is highly recommended.

## Hardware Connections
**NI DAQ USB Board:** Data is routed from the load cells to amplifiers with +/-5V inputs which are then routed to specific input channels on the NI DAQ board. If collecting EMG data, data is routed from the wireless sensor to the Delsys Trigno EMG Base Station, and then to the NI Board via a 64-pin cable adapter provided by Delsys. The NI DAQ device (and EMG Base Station if using) must be connected to a USB-A port on the computer. 

**Accelerometer:** The Phidgets accelerometer device must be connected to the USB-A port on the computer.

> [!IMPORTANT]
> To collect data from the load cell, EMG base, and accelerometer simultaneously, this requires 3 _separate_ USB-A ports on the computer. Adapters will not work as a substitute. 

**Secondary display:** In our experiments, subjects were provided with the Lissajous feedback on a head mounted display via HDMI cable. This can be substituted for another display (e.g., monitor, projector, etc.) if desired. Make sure to select "Extend these displays" once the extra display is connected.

## Usage

<img width="2144" height="988" alt="LabVIEW GUI front panel" src="https://github.com/user-attachments/assets/b7766d84-67cc-4a16-8582-440916b38bfc" />

**Fig. 1: Main GUI front panel**


1. If collecting EMG data:
   - Open the Trigno Control Utility app.
   - Pair EMG sensors.
   - For each sensory, click the gear icon and verify in the Sensor Configuration window that the "EMG + Accelerometer" option is selected.
   - Select "Start Analog".
> [!IMPORTANT]
> If this "Start Analog" is not selected before proceeding, EMG data will not be recorded.
2. Open the project _BimCoord DAQ.lvproj_. Right click and Run the _Launcher.vi_.
3. Channel Config (Fig. 2)
   - Use the drop down menus to select channels from the NI DAQ board.
   - If only EMG data will be collected and no force data, check the "Use EMG Master Mode" box. This will only use the channels in the EMG Master Mode box.
> [!NOTE]
> Ensure that the device number matches the device that is connected to the computer. Only valid device numbers will appear in the drop down menu.\
> Ensure that the same DAQ physical channel is not selected for more than one drop down menu. Channels may overlap between the upper channel group (force, EMG, Potentiometer) and the lower EMG Master Mode channel group.

<img width="320" height="290" alt="Channel config tab" src="https://github.com/user-attachments/assets/908c02db-8057-42cc-9444-fa087b0f7e7e" />

**Fig. 2: Channel Config tab**

4. Calibration (Fig. 3)
   - Volts to Newtons Regression: During initial setup with a load cell/transducer, a regression equation must be calculated to convert volts to newtons. This can be done by placing weights of known mass onto the transducer and selecting "Tare Transducers" and then recording the weight and voltages in the 3x3 table. Once this is completed for 3 masses, select "Calculate Regression". The regression equation (y = mx + b) will be displayed under "Regression in Use".
   - Zero out load cells by selecting "Tare Transducers".

<img width="320" height="290" alt="Devices tab" src="https://github.com/user-attachments/assets/afc2622e-761e-4b98-93d7-a5ca5afd70aa" />

**Fig. 3: Devices tab**

5. MVC data collection (Fig. 4)
   - Use this tab to collect maximum voluntary contraction (MVC) from participants.
   - This is used to scale the Lissajous feedback plot to prevent subject fatigue. 20% MVC is typically used, although the sliding bar may be used to calculate other percentages if desired.
   - Data will be saved to an excel file at the path and name specified (Fig. 1A).

<img width="320" height="290" alt="MVC tab" src="https://github.com/user-attachments/assets/0b8834e9-ef5c-4665-a582-38ad58393611" />

**Fig 4: Maximum Voluntary Contraction (MVC) tab**

6. Task Configuration
   - Each task is composed of a series of trials. Old task configurations may be loaded or new ones saved using the drop down in the top of Fig. 1B.
> [!NOTE]
> Acceleration sampling rate cannot be set to any value under 5 Hz. 
   - There are two modes: manual and in-flight.
     - Manual mode (Fig. 1B): trials proceed sequentially after predefined intervals ("Time Between Trials (s)". Use the bottom box to specify the number of trials, time between trials, and the bimanual coordination pattern. This mode only allows for one coordination pattern to be used per task (set of trials).
     - In-flight mode: data collection pauses after each trial and resumes upon user input. Check "In-flight?" in the "Parabolic Flight" tab (Fig. 5) to activate this mode. When enabled, parameters specified in the Parabolic Flight tab (Fig. 4C), including the number of parabolas/trials and coordination patterns, override corresponding values in the "Task Config" tab.
       
<img width="200" height="215" alt="Parabolic flight tab" src="https://github.com/user-attachments/assets/8ebde2b0-2065-46d9-ae47-f30e14966870" />

**Fig. 5: Parabolic Flight tab**
       
   - For both modes: Use the fields in the top box to specify task type, input type, toggle on/off EMG and acceleration data collection, and sampling rate for EMG and acceleration. Force sampling rate, trial duration, and adjusted MVC are specified in the bottom box.
   - The "Preview Template" button can be used to view a preview of the coordination pattern currently specified in the bottom box of the "Task Config" tab.

7. Data Collection
   - If desired, check the box "Display task on second monitor" (Fig. 1E). This will display only the Lissajous plot on your secondary display.
   - Set the path and file name (Fig. 1A). A prompt will warn you if a file with that name already exists.
   - Once the task is configured as desired, select "Start Trials" to launch the task. Live data streams will appear for the NI DAQ Board and accelerometer if applicable (Fig. 1C) and a copy of the Lissajous display will appear in the lower right corner (Fig. 1E).
   - User Markers: Use the "Acceleration Markers" tab (Fig. 6) to record time-stamped event markers as desired during data collection. Users may manually click the marker buttons or press the corresponding key noted in parentheses. The marker name and timestamp will be recorded in a separate sheet the final excel file.
   
<img width="200" height="215" alt="Acceleration markers tab" src="https://github.com/user-attachments/assets/1165bbbb-099f-474e-b149-b13ee0522559" />

**Fig. 6: Acceleration Markers tab**

> [!CAUTION]
> If at any time the trial needs to be stopped, use the red "Abort Task" button in the bottom left of the "Task Config" tab. Stopping the task any other way can result in "orphaned" actors running the background and may require a full project or LabVIEW restart to resolve.

  - Once the task concludes, the _SaveAsExcel.vbs_ script will consolidate data written to the temporary TDMS files into one Excel file.

> [!IMPORTANT]
> During this process, the script will open Excel in the background for a few seconds. Do NOT close Excel. It will close on its own.

> [!IMPORTANT]
> Check if the final Excel file is saved and the expected data is present. If the Excel file is not there, likely causes include the TDMS Excel Add-In not activating or insufficient time for the .vbs script to run. Check if the TDMS add in works by trying to open a TDMS file. If the Excel Add-In did not have sufficient time to run, this may be an issue with file size (high sampling rate, many trials, long trial duration) or low computing power available.

> [!TIP]
> If the Excel sheet was not saved successfully, data recovery may be possible!
> - First, check if the TDMS file was saved to the folder specified. These can easily be converted to an Excel file using the TDMS Excel Add-In by opening the TDMS file. If the Add-In is functioning correctly, it will open in a new Excel workbook which can then be saved.
> - If the TDMS file was not saved to the folder specified, locate the _temp_ni.tdms_, _temp_ni.tdms_index_, _temp_accel.tdms_, and _temp_accel.tdms_index_ files in the main project folder. These files contain the raw data that was written during acquisition. Copy these files to another location and rename them. 

