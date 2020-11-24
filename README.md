# obf_autop
Version 3.8<br>
AutOP is a Python script that creates a report (Operat) out of downloaded raw SAP data.

## Install

requirements:<br>
python 3.8.2<br>
numpy-1.18.1<br>
pandas-1.0.3<br>
matplotlib-3.0.1<br>
openpyxl-3.1.3<br>
xlrd-1.2.0<br>
python-docx-0.8.10 (install with pip)<br>
pyqt-5.9.2 (optional for GUI)<br>

### Install instructions
The easiest way to install all required libraries for running the AutOP code, is by installing conda first and creating a conda environment.

1. download and install conda from:
  * Anaconda (recomended):  https://www.anaconda.com/products/individual<br>
or from:<br>
  * Miniconda :             https://repo.anaconda.com/miniconda/<br>
  
2. create an environment<br>
```conda create --autop myenv```

3. activate the newly created enviroment<br>
```conda create --autop myenv```




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
