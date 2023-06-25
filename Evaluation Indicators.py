import pandas as pd
from pandas import DataFrame
from sklearn.metrics import r2_score, mean_absolute_error


# Convert to pkl format file
def to_pkl1(path):
    # Read excel files
    df = DataFrame(pd.read_excel(path))
    df.to_pickle('Spectral Data/Original spectral data/TL_Data1/Combined_gas_data_vmd.pkl')


# Convert to pkl format file
def to_pkl2(path):
    # Read excel files
    df = DataFrame(pd.read_excel(path))
    df.to_pickle('Spectral Data/Original spectral data/TL_Data2/Mixed_gas_data_vmd.pkl')


def read_data(path):
    df = pd.read_excel(path, header=None)
    pre_y = df.values[:, :2]
    test_y = df.values[:, -2:]
    return pre_y, test_y


def score_R_2(pre_y, test_y):
    CS2_pre_y = pre_y[:, 0]
    SO2_pre_y = pre_y[:, 1]
    CS2_test_y = test_y[:, 0]
    SO2_test_y = test_y[:, 1]
    CS2_r2_score = r2_score(CS2_pre_y, CS2_test_y)
    SO2_r2_score = r2_score(SO2_pre_y, SO2_test_y)
    print("CS2_r2_score:", CS2_r2_score)
    print("SO2_r2_score:", SO2_r2_score)


def score_MAE(pre_y, test_y):
    CS2_pre_y = pre_y[:, 0]
    SO2_pre_y = pre_y[:, 1]
    CS2_test_y = test_y[:, 0]
    SO2_test_y = test_y[:, 1]

    CS2_MAE = mean_absolute_error(CS2_pre_y, CS2_test_y)
    SO2_MAE = mean_absolute_error(SO2_pre_y, SO2_test_y)
    print("CS2_MAE:", CS2_MAE)
    print("SO2_MAE:", SO2_MAE)


if __name__ == '__main__':
    '''
    The program takes a short while to run depending on the configuration of the computer
    '''
    evaluation = False

    if evaluation:
        path = ""  # Path of the document to be evaluated
        pre_y, test_y = read_data(path)
        score_MAE(pre_y, test_y)
        score_R_2(pre_y, test_y)
    else:
        path1 = "Spectral Data/Original spectral data/TL_Data1/Combined_gas_data_vmd.xlsx"
        # Convert to pkl format file
        to_pkl1(path1)
        path2 = "Spectral Data/Original spectral data/TL_Data2/Mixed_gas_data_vmd.xlsx"
        # Convert to pkl format file
        to_pkl2(path2)
