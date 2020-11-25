# obf_autop
Version 3.8<br>
AutOP is a Python script that creates a report (Operat) out of downloaded raw SAP data.

## Install

### requirements:<br>
python 3.8.5<br>
numpy 1.19.2<br>
pandas 1.1.3<br>
matplotlib 3.3.2<br>
openpyxl 3.0.5<br>
xlrd 1.2.0<br>
python-docx 0.8.10<br>
pyqt 5.9.2 (optional for GUI)<br>

### Install enviroment
Although it is not strictly necessary to create an enviroment with conda, it is highliy recomended to use conda to assure no problems with already installed package versions. If you have conda (Anaconda or Miniconda) already installed please skip point 1. and continue with 2. If you are unsure if you have conda installed you can use the following command to check:<br>
```which conda```

1. Download and install conda from:<br>
  * Anaconda (recommended): https://www.anaconda.com/products/individual<br>
or from:<br>
  * Miniconda :             https://repo.anaconda.com/miniconda/<br>
  
2. Create an environment for AutOP<br>
```conda create --name autop```

3. Activate the newly created enviroment<br>
```conda activate autop```

4. Install packages for AutOP<br>
```conda install python=3.8.5 numpy=1.19.2 pandas=1.1.3 matplotlib=3.3.2 openpyxl=3.0.5 xlrd-1.2.0 ipykernel```

5. Install python-docx package<br>
```pip install python-docx```

6. Add kernel to jupyter notebook<br>
```python -m ipykernel install --user --name autop```

7. Deactivate kernel<br>
```conda deactivate```

## Copy repository

1. Navigate to home directory<br>
```cd```

2. Make a new directory and navigate to it<br>
```mkdir -p Code/python & cd Code/python```

3. Clone the github enviroment (or download the zipped version)<br>
```git clone https://github.com/satlawa/obf_autop.git```

## Prepare Data
Download all nescesery data from SAP and put the files in the appropiate folder.<br>
The folder should have following naming convention: **TOXXXX** where **XXXX** stands for the "Teiloperats-ID" (for example "**TO1284**").<br>
To create the full report the following data has to be moved into the folder (assuming example "**TO1284**"):<br>

Folder -> TO1284<br>
* to_1284_bw_alter.xlsx
* to_1284_bw_neig.xlsx
* to_1284_bw_seeh.xlsx
* to_1284_bw_uz.xlsx
* to_1284_bw_wp.xlsx
* to_1284_bw_zufaellige.xlsx
* to_1284_hs_bilanz.XLS
* to_1284_hs_bilanz_old.XLS
* to_1284_sap.XLS
* to_1284_sap_natur.XLS

Folder -> stichprobe<br>
* SPI_2019.txt

Folder -> klima<br>
* 1_Lufttemperatur.txt
* 2_Niederschlag.txt
* 4_Schnee.txt

Folder -> dict<br>
* several static data files

## Run AutOP
AutOP can be run within a jupyter notebook or with a light weight GUI.

### Run with Jupyter Notebook:

1. Start jupyter notebook<br>
```jupyter notebook```

2. Usually if the enviroment is called **autop** the right kernel should be picked automatically. In any other case please make sure to use the right kernel by checking the name in the upper right corner. In case a diffrent kernel is active please change to the autop kernel by selecting upper tab<br>
**Kernel -> Change kernel -> autop**<br>

2. Navigate to obf_autop folder and open<br>
**run_AutOP.ipynb**<br>

3. Set the data path by setting the variable<br>
**data_path**<br>

4. select all sections of the report that you want to create<br>
* if set to **1** -> section will be created<br>
* if set to **0** -> section will NOT be created<br>

### Run with GUI:

activate enviroment<br>
```conda activate autop```<br>
<br>
navigate to the obf_autop folder<br>
```cd Code/python/obf_autop```
<br>
start AutOP with command<br>
```python AutOP.py```<br>
