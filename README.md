# obf_autop
Python script that creates a report out of SAP data.

## Install

requirements:<br>
python 3.7.1<br>
numpy-1.15.4<br>
pandas-0.23.4<br>
matplotlib-3.0.1<br>
docx (install with pip)<br>
pyqt-5.9.2 (optional for GUI)<br>

### Install instructions
The easiest way to install all required libraries for running the AutOP code, is by installing conda first and creating a conda environment.

## Prepare Data
Download all nescesery data from SAP and put the files in the appropiate folder. Detailed instructions are provided in the data folder.

## Running AutOP
AutOP can be run with a light weight GUI or directly from a jupyter notebook.

### Run AutOP with GUI:

activate enviroment<br>
$ conda activate autop<br>
<br>
navigate to the obf_autop folder<br>
$ cd code/obf_autop<br>
<br>
start AutOP with command<br>
$ python AutOP.py<br>

### Run AutOP from jupyter notebook:

open "jupyter notebook main_03.ipynb" and choose the right kernel.
