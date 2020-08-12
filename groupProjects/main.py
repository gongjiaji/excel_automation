import pandas as pd

bigExcel = pd.read_excel("big.xlsx")
bigExcel['projectName'] = bigExcel["Ref."].apply(lambda x: "-".join(str(x).split("/")[:-1]))
groupByProjectName = bigExcel.groupby("projectName")

for name, group in groupByProjectName:
    group: pd.DataFrame = group
    group.rename(columns={"Ref.": "PO NO.","Doc. No.":"Inv No."}, inplace=True)
    del group["projectName"]
    group.sort_values(by='Doc. No.')
    outstanding = group["Outstanding"].sum()
    print(len(group))

    outDict = {"Outstanding": outstanding}
    group = group.append(outDict, ignore_index=True)
    group.to_excel(name+"-"+str(len(group))+".xlsx", index=None)
