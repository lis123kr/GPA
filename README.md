# GPA - Genetic Polymorphism Analyzer

## Introduction

GPA consists of analyzer and visualization for genetic polymorphism of DNA sequences.
The analyzer compares MAF(Minor Allele Frequency) between each sequences, finds genetic polymorphism and reports it (to Excel format) according to [genome structure, repeat region, ORF, NCR, Major/Minor combinations].
After finishing analyzer, the visualization will show up. You can easily check the changes of MAF at each base pair (position) and interact with it by clicks, drags on it for more detailed information.

GPA is implemented in plain python(v3.6), and available as open-source under the flexible Apache License 2.0.

# Prerequisities

You will need:

* Python3 (over v3.6)
* if pip isn't installed, try `python get-pip.py`
* `pip install -r requirements.txt`

## Run
and you can run it:

* `python GPA_UI.py`

if a graphic user interface showed up. then,

* load a file(an Excel file, maybe take few minutes)
* click sheets to analyze
* insert information of the file
* choose analyzation type and `OK` 

then, it will run.

### GUI example
![Image](asset/GUI.PNG =450x540)

### Visualization examples
![Image](asset/visualization.PNG =500x350)
![Image](asset/visualization-3.JPG =500x350)


### License

GPA is licensed under the Apache License, Version 2.0.
Each source code file has full license information.
