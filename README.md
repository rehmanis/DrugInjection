# DrugInjection

### Introduction
The DrugInjection repository was created for use by the UBC BioMEMS group. It serves to automate injection of drugs contained on a 96-well microplate into microfluidic devices. The apparatus used to achieve this are as follows:

1. Zaber T-LSR450D Linear Stage (X-Axis stage)
2. Zaber T-LSR150B Linear Stage (Y-Axis stage)
3. Melles Griot 17NST101 Linear Stage (Z-Axis Stage)
4. KDS 230 Infusion/Withdrawal Syringe Pump

The contents of the repository is detailed in the following section.
***
### Contents of Repository
The contents of the repository allow the user to define the experiment through a macro-enabled excel file and to run the code through the DrugInjection GUI (which should be published locally on the system for quick access). 

Folder | Contents | Description 
--- | --- | --- 
AutomatedDrugInjection | <ol><li>C# Solution for Stage/Pump automation</li><li>Excel template for Experiment</li></ol> | The folder contains the Visual Studio C# solution for automating the linear stages and syringe pump to perform an experiment. It also contains the 3rd party libraries (dlls) required for operation. The Excel file is a macro-enabled template to allow users to easily define experiments (see Defining the Experiment)
DrugInjection GUI | C# solution for GUI | The folder contains the Visual Studio C# Solution for a graphical user interface to allow users to easily load and run their experiments. The GUI was developed using the C# WinForms framework.
Equipment Manuals | Manuals for Stages/Pump | The folder contains manuals for operation of the Melles Griot Linear Stage and the KDS230 Syringe pump.
Excel VBA | (Ignore) | (Ignore)
Time Calculation | C# solution for time calculation | The folder contains the Visual Studio C# Solution for calculating the approximate running time of the experiment. This is used by the GUI to update its progress bar.
***
### Defining the Experiment
The macro-enabled Excel file ([Excel Definition Template](AutomatedDrugInjection/automatedDrugTestTemplate.xlsm)) is used to define the experiment. It has been programmed through the developer-mode Visual Basic environment in Microsoft Excel 2010. The following sections outline the use of the template, structure of the code, functions available for experiment definition, the use of color indicators to guide user input.
####Template Structure
The experiment definition requires the use of Sheets 1 and 2 of the Excel workbook template. Sheet 1 requires the user to specify functions and parameters to be executed as part of the experiment. Details of the positions of the drugs contained in the microplate are contained in Sheet 2. The figure below shows a sample experiment defined through the template:
***
### Running the Experiment
#### Calibration
