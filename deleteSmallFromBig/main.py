import pandas as pd
import os

# file process
allfiles = []
for _, _, filenames in os.walk("./"):
    allfiles.append(filenames)
a = []
for x in allfiles:
    a += x
lastResults = [x for x in a if "Result_" in x]
for x in lastResults:
    os.remove("./" + x)
keyword: str = [x for x in a if "keyword" in x][0].split("!@#$%")[0]

# get excel file
big = [x for x in set(a).difference(set(lastResults)) if "big" in x]
smallExcel = pd.concat(
    [pd.read_excel(x, converters={keyword: str}) for x in a if "small" in x]
)
# Get result
for b in big:
    bigExcel: pd.DataFrame = pd.read_excel(b)
    # 如果小表的物品在大表里出现了, 则在大表里把这一行删掉
    smallCodes = [x for x in smallExcel[keyword]]
    resultExcel = bigExcel[~bigExcel[keyword].isin(smallCodes)]
    resultExcel.to_excel("Result_" + b, index=False)
