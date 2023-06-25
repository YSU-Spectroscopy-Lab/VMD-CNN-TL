# UV-DOAS-VMD-CNN-TL

### Catalogue
* Model Structure and Parameters (directory)
* Results Data (directory)
* Run_data (directory)
* Spectral Data (directory)
* Data Process.py (python file)
* Evaluation Indicators.py (python file)
* Transfer Learning.py (python file)
* VMD Process.py (python file)
* vmdpy.py (python file)
* requirements.txt (TXT)

### Content and functions
> Model Structure and Parameters: This directory holds the structure and parameters of the trained model for the project.

> Results Data: This directory holds the final predictions of all models for the project.

> Run_data: The purpose of this directory is to hold the cache files for the project runs.

> Spectral Data: This directory contains the raw spectral data, the VMD pre-processed data, the test data for the project. Also, field test data is included.

> Data Process.py: Reading and processing data.

> VMD Process.py: VMD pre-processing of the data.

> vmdpy.py: VMD function package.

> Evaluation Indicators.py: Two functions, one is the conversion of .xlsx files into .pkl files and the other is the evaluation of the generated model prediction data.

> Transfer Learning.py: Training multilayer 1D convolutional neural networks with migration learning.

### Environment
- TensorFlow-gpu==2.4.0
- cudnn==8.0.5.39
- cudatoolkit==11.0.3
- The package versions of project are in requirements.txt.

### Run
1. First run **Data Process.py**, the details of which are commented in that file.
2. When you have enough of the combined gas and mixture processed file, run **VMD Process.py**. Specific operational details and notes are annotated within this file.
3. Once you have the VMD-processed file, run the built-in function topkl() for **Evaluation Indicators.py** to convert the .xlsx file into a .pkl file.
4. Once you get the .pkl file, you can run the **Transfer Learning.py**. Training of the model is carried out and specific details of the operation are provided in the comments within the file.
5. The final predictions are evaluated by means of the **Evaluation Indicators.py** mainland evaluation function.
