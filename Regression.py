import pandas as pd
import statsmodels.api as sm
from io import StringIO

data = pd.read_excel(open('X and Y variablesV2.xlsx', 'rb'), sheet_name='Variables Filtered yrs stacked') 


# Raw regression, every label and every coef. Writes to excel file
def regress(filename, Xs, Ys):
    labelCount = {}
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    for i in range(len(Ys)):
        label = Ys[i]
        labelX = Xs[i]
        dataX = data[labelX]
        dataX = dataX.fillna(dataX.median())
        dataY = data.loc[:,label]
        dataY = dataY.fillna(dataY.median())
        X = sm.add_constant(dataX.values)
        ols1 = sm.OLS(dataY.values, X)
        ols1Result = ols1.fit()
        summary = ols1Result.summary2(xname=['const']+labelX)
        result = summary.tables[0].append(summary.tables[1].reset_index(), ignore_index=True, sort=False)
        if (label in labelCount):
            summary.tables[0].to_excel(writer, sheet_name=f'{label[:15]}{labelCount[label]} Results')
            summary.tables[1].to_excel(writer, sheet_name=f'{label[:15]}{labelCount[label]} Coef.')
            labelCount[label] += 1
        else:
            summary.tables[0].to_excel(writer, sheet_name=f'{label[:15]} Results')
            summary.tables[1].to_excel(writer, sheet_name=f'{label[:15]} Coef.')
            labelCount[label] = 1
    writer.save()


# Remove each non-statistically significant coef. until either all coef. p-value are under 0.05 or 
# there is only one coef. left. Writes to Excel file
def refineRegress(filename, Xs, Ys):
    labelCount = {}
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    for i in range(len(Ys)):
        label = Ys[i]
#         print(label)
        labelX = Xs[i]
        dataY = data.loc[:,label]
        dataY = dataY.fillna(dataY.median())
        run = True
        while(run):
            run = False
            dataX = data[labelX]
            dataX = dataX.fillna(dataX.median())
            X = sm.add_constant(dataX.values)
            ols1 = sm.OLS(dataY.values, X)
            ols1Result = ols1.fit()
            vals = ols1Result.summary2(xname=['const']+labelX).tables[1]
            for j in range(len(labelX)):
                if (vals['P>|t|'][j+1] > 0.05 and len(labelX) > 1):
#                     print(vals['P>|t|'][j+1])
#                     print('removing', labelX[j])
                    labelX.pop(j)
                    run = True
                    break
#         print('final', labelX)
        dataX = data[labelX]
        dataX = dataX.fillna(dataX.median())
        X = sm.add_constant(dataX.values)
        ols1 = sm.OLS(dataY.values, X)
        ols1Result = ols1.fit()
        summary = ols1Result.summary2(xname=['const']+labelX)
        result = summary.tables[0].append(summary.tables[1].reset_index(), ignore_index=True, sort=False)
        if (label in labelCount):
            summary.tables[0].to_excel(writer, sheet_name=f'{label[:15]}{labelCount[label]+1} Results')
            summary.tables[1].to_excel(writer, sheet_name=f'{label[:15]}{labelCount[label]+1} Coef.')
            labelCount[label] += 1
        else:
            summary.tables[0].to_excel(writer, sheet_name=f'{label[:15]} Results')
            summary.tables[1].to_excel(writer, sheet_name=f'{label[:15]} Coef.')
            labelCount[label] = 1
#         print()
#         print()
#         print()
    writer.save()


# All possible X and Y [will need to be checked]
ynames = ['DARTFreq', 'AutoFreq', 'Tier3Freq', 
          'LostTimeFreq',  'NDPPH', 'Mileage Index (% to Plan)',
          'Inside PPH', 'OA per TSP (% to Plan)']

xnames = ['CDEV', 'GRW', 'MGC', 'MGR', 'ERI', 'CMGT', 'COMP', 
          'ENG', 'ECI', 'AUTO', 'SFT', 'Overall', 'QPR Score', 
         'Urban', 'Suburban', 'Rural', 'Super Rural',
         'HIS', 'W', 'B', 'AA', 'Male (1) / Female (0)', 'Age', 'Tenure']

# Labels `Ys` and corrisponding coef. `Xs`. They need to be in order relative to each other.
Ys = ['DARTFreq',
    'AutoFreq',
    'LostTimeFreq',
    'NDPPH',
    'Mileage Index (% to Plan)',
    'Inside PPH',
    'OA per TSP (% to Plan)',
    'AutoFreq',
    'Urban',
    'Suburban',
    'Rural',
    'Super Rural']

Xs = [
    xnames.copy(),
    xnames.copy(),
    xnames.copy(),
    xnames.copy(),
    xnames.copy(),
    xnames.copy(),
    xnames.copy(),
    ['NDPPH'],
    ['AutoFreq'],
    ['AutoFreq'],
    ['AutoFreq'],
    ['AutoFreq'],
]

regress('Regression Raw.xlsx', Xs.copy(), Ys.copy())
refineRegress('Regression Refined.xlsx', Xs.copy(), Ys.copy())
print('DONE!!!') 
