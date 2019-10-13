import pyexcel as pe
import datetime
import calendar
import numpy as np
import pandas as pd

def dataset_date_clean(data):
    for i in np.arange(len(data)).tolist():
        if data[i, Monthly_Date] == "":
            data = np.delete(data, i, 0)

def dictionary_creator(data, dictionary):
    for i in np.arange(1, len(data)).tolist():
        dictionary[str(data[i, Investment_Options].item()), date_converter(data[i, Monthly_Date])] = i

def array_equaliser(output1, output2):
    biggest_array = max(len(output1), len(output2))
    smallest_array = min(len(output1), len(output2))
    for i in np.arange(biggest_array - smallest_array).tolist():
        if len(output1) < len(output2):
            output1.append("")
        else:
            output2.append("")

def problemdatefinder(num_months, options):
    for i in np.arange(len(options) - 2).tolist():
        loop_status = True
        try:
            for j in np.arange(num_months).tolist():
                index = 1.0
                if not loop_status:
                    break
                for k in range(3):
                    index = index * (1.0 + float(dataset[datasetdict[options[i + 2], dates[j + k]], One_Month]))
                if abs((index - 1.0) - float(dataset[datasetdict[options[i + 2], dates[j]], Three_Month])) > (bound/100):
                    for l in np.arange(j + 1, num_months).tolist():
                        index = 1.0
                        for m in range(3):
                            index = index * (1.0 + float(dataset[datasetdict[options[i + 2], dates[l + m]], One_Month]))
                        if abs((index - 1.0) - float(dataset[datasetdict[options[i + 2], dates[l]], Three_Month])) < (bound/100):
                            options[i + 2] = options[i + 2] + " - " + dates[l - 1]
                            loop_status = False
                            break
                        elif j == num_months:
                            options[i + 2] = options[i + 2] + " - All data is problematic"
        except(TypeError):
            options[i + 2] = options[i + 2] + " - Inconclusive"

def date_converter(Date):
    ts = pd.to_datetime(Date)
    d = ts.strftime("%Y-%m-%d")
    return str(d[0].item())

def index_creator(num_months, Option, period, output):
    index = 1.0
    for i in np.arange(num_months).tolist():
        index = index * (1.0 + float(dataset[datasetdict[Option, dates[i]], One_Month]))
    if (bound/100) < ((index - 1.0) - float(dataset[datasetdict[Option, dates[0]], period])) or ((index - 1) - float(dataset[datasetdict[Option, dates[0]], period])) < -(bound/100):
        output.append(Option)

def missingdata_report(num_months, missingdata):
    if not missingdata:
        print('\nThere are no options with missing data for the', num_months, 'month period!\n')
    else:
        print('\nThe following options don\'t possess', num_months, 'months of data to analyse:\n')
        for i in np.arange(len(missingdata)).tolist():
            print('\t', missingdata[i])

def problemdata_report(num_months, problemdata):
    if not problemdata:
        print('\nThere are no options whose performance is outside the bounds of', str(bound) + '%','and', str(-bound) + '%'' for the', num_months,'month period!\n')
    else:
        print('\nThe following options are outside the specified bounds of', str(bound) + '%','and', str(-bound) + '%'' for the', num_months,'month period:\n')
        for i in np.arange(len(problemdata)).tolist():
            print('\t', problemdata[i])

def Output_Summary_Creator(array, output):
    for i in np.arange(len(output)).tolist():
        x = [output[i]]
        array = np.append(array, x)
    return array

fname = "SeptemberTest.xlsx"
# input('What is the name of the file, including its extension:')

ddate = input('What date is the data effective for? (YYYY-MM-DD):')
bound = float(input('What bounds of difference would you like to test for?: '))

dataset = pe.get_array(file_name=fname, skip_empty_rows=True)
dataset = np.array(dataset)

# As the export file has two headers, neither of which can be used on their own,
# this process combines the two into one.
for i in np.arange(len(dataset[0]) - 1).tolist():
    dataset[0, i] = dataset[0, i] + ' ' + dataset [1, i]

dataset = np.delete(dataset, 1, 0)

print('Processing...')

# This searches for the headers of the columns which will be used and returns
# the values to make it easier for the program to ssearch for them.
Investment_Options = np.where(dataset == 'Investment Options Name')[1]
One_Month = np.where(dataset == 'Monthly % - 1 Month')[1]
Three_Month = np.where(dataset == 'Monthly % - 3 Month')[1]
One_Year = np.where(dataset == 'Monthly % - 1 Year')[1]
Monthly_Date = np.where(dataset == "Monthly Date")[1]

# dataset_date_clean(dataset)

datasetdict = dict()

dictionary_creator(dataset, datasetdict)

# This line sets the date for which the performance is to be checked.
ddate = datetime.datetime.strptime(ddate, "%Y-%m-%d")

# Uses the input date to create a numpy list of it and the previous 11 months
# which will the basis for investigation.
years = dict()
months = dict()
days = dict()
for i in np.arange(12).tolist():
    if (ddate.month - i) <= 0:
        months[i] = 12 + ddate.month - i
        years[i] = ddate.year - 1
    else:
        months[i] = ddate.month - i
        years[i] = ddate.year
for i in np.arange(12).tolist():
    days[i] = calendar.monthrange(years[i], months[i])[1]

dates = np.zeros(0)
for i in np.arange(12).tolist():
    if months[i] < 10:
        months[i] = '0' + str(months[i])
        dates = np.append(dates, (str(years[i]) + '-' + str(months[i]) + '-' + str(days[i])))
    else:
        dates = np.append(dates, (str(years[i]) + '-' + str(months[i]) + '-' + str(days[i])))

for i in np.arange(12).tolist():
    dates = np.append(dates, dates[i])

# Creation of unique list of options to analyse.
OptionList = []
for i in np.arange(2, (len(dataset))).tolist():
    OptionList.append(dataset[i, 5])
OptionList = set(OptionList)

missingdata3m = []
missingdata12m = []
problemdata3m = ["The following is problematic 3 month data:", ""]
problemdata12m = ["The following is problematic 12 month data:", ""]
output_list = ["Superannuation Performance Review Summary"], ["Effective Date: " + ddate.strftime("%d/%m/%Y")], ["Bounds of error: " + str(bound) + "%" + " and " + str(-bound) + "%"], [""]
output_array = np.array(output_list)
output_list2 = [""], [""], [""], [""]
output_array2 = np.array(output_list2)

for i in OptionList:
    try:
        index_creator(3, i, Three_Month, problemdata3m)
    except (TypeError, KeyError):
        if i not in missingdata3m:
            missingdata3m.append(i)

for i in OptionList:
    try:
        index_creator(12, i, One_Year, problemdata12m)
    except (TypeError, KeyError):
        if i not in missingdata12m:
            missingdata12m.append(i)

problemdata3m[2:] = sorted(problemdata3m[2:])
problemdata12m[2:] = sorted(problemdata12m[2:])

problemdatefinder(3, problemdata3m)
problemdatefinder(12, problemdata12m)

array_equaliser(problemdata3m, problemdata12m)

output_array = Output_Summary_Creator(output_array, problemdata3m)
output_array2 = Output_Summary_Creator(output_array2, problemdata12m)
output_array_final = np.column_stack((output_array, output_array2))

df = pd.DataFrame(output_array_final)
df.to_csv("Performance_Test_" + ddate.strftime("%Y%m%d") + ".csv", index = False, header = False)

print('Your file is ready!')
