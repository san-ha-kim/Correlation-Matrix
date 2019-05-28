import pandas as pd
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
from pandas import DataFrame as df
from scipy import stats


def calculate_pvalues(df):
    df = df.dropna()._get_numeric_data()
    dfcols = pd.DataFrame(columns=df.columns)
    pvalues = dfcols.transpose().join(dfcols, how='outer')
    for r in df.columns:
        for c in df.columns:
            pvalues[r][c] = round(stats.pearsonr(df[r], df[c])[1], 4)
    return pvalues


filepath = "" #LOCAL FILE PATH

# -- Import excel file; requires 'r' prefix to address special character strings "\"

dropped_masterfile_withPID = pd.read_excel(filepath, sheet_name="droppedPosner")

# -- Remove irrelevant columns
dropped_masterfile = dropped_masterfile_withPID.drop(columns=["PID", "AFDQ Total", "AFDQ SD", "AISS Total", "AISS SD","CFQ Total", "CFQ SD", "DBQ Total", "DBQ SD","Posner n(valid, correct)","Posner n(invalid, correct)","MOT Median", "MOT Mode"])


# -- Convert to pandas data frame and change column names
df_masterfile = df(dropped_masterfile)

df_masterfile.columns = ["AFDQ %", "AFDQ mean", "AISS %", "AISS mean", "CFQ %", "CFQ mean",
                         "DBQ %", "DBQ mean", "SC(V) %", "SC(I) %", "SC(V) mean RT", "SC(I) mean RT", "MOT mean score", "MOT percentage score"]

# -- Clean the data by removing NaNs
df_masterfile= df_masterfile.dropna()
# print(df_masterfile)

# -- Standardise using z-scores and convert to data frame
masterfile_z = stats.zscore(df_masterfile, axis=0, ddof=0)
df_masterfile_z = df(masterfile_z)

# Generate correlation matrix; round to 3 decimal places

df_masterfile_z.columns = ["AFDQ %", "AFDQ mean", "AISS %", "AISS mean", "CFQ %", "CFQ mean", "DBQ %", "DBQ mean", "SC(V) %", "SC(I) %", "SC(V) mean RT", "SC(I) mean RT", "MOT mean score", "MOT percentage score"]

masterfile_pearson_matrix = df_masterfile_z.corr(method="pearson")
masterfile_pearson_matrix = masterfile_pearson_matrix.round(3)

masterfile_pearson_not_z = df_masterfile.corr(method="pearson")
masterfile_pearson_not_z = masterfile_pearson_not_z.round(3)
mfpnz = masterfile_pearson_not_z

# Generate correlation matrices with p-value significance
masterfile_pearson_pv = calculate_pvalues(masterfile_pearson_matrix)
df_masterfile_pv = df(masterfile_pearson_pv)
df_masterfile_pv = df_masterfile_pv.round(3)

masterfile_NZ_pv = calculate_pvalues(masterfile_pearson_not_z)
df_mnzpv = df(masterfile_NZ_pv)
df_mnzpv = df_mnzpv.round(3)

# -- Apply the p-val significance mask
pv1 = masterfile_pearson_matrix.applymap(lambda x: "{}*".format(x))
pv2 = masterfile_pearson_matrix.applymap(lambda x: "{}**".format(x))
pv3 = masterfile_pearson_matrix.applymap(lambda x: "{}***".format(x))

masterfile_pearson_matrix = masterfile_pearson_matrix.mask(masterfile_pearson_pv <= 0.1, pv1)
masterfile_pearson_matrix = masterfile_pearson_matrix.mask(masterfile_pearson_pv <= 0.05, pv2)
masterfile_pearson_matrix = masterfile_pearson_matrix.mask(masterfile_pearson_pv <= 0.01, pv3)

nz1 = mfpnz.applymap(lambda x: "{}*".format(x))
nz2 = mfpnz.applymap(lambda x: "{}**".format(x))
nz3 = mfpnz.applymap(lambda x: "{}***".format(x))

mfpnz = mfpnz.mask(masterfile_pearson_pv <= 0.1, pv1)
mfpnz = mfpnz.mask(masterfile_pearson_pv <= 0.05, pv2)
mfpnz = mfpnz.mask(masterfile_pearson_pv <= 0.01, pv3)

# -- Change the data frame column names
df_masterfile_pearson_matrix = df(masterfile_pearson_matrix)
df_mfpnz = df(mfpnz)

# -- Export the results to excel files
pearson_path = r"C:\Users\USER\Dropbox\MEng\Project\Experiments\OUTPUT\full corr matrix pearson standardised.xlsx"
pearson_writer = pd.ExcelWriter(pearson_path, engine="xlsxwriter")

df_masterfile_pearson_matrix.to_excel(pearson_writer, "corr")
df_masterfile_pv.to_excel(pearson_writer, sheet_name="p values")
df_mfpnz.to_excel(pearson_writer, sheet_name="not standardised")

pearson_writer.save()
pearson_writer.close()
