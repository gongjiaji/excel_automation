import pandas as pd
import rich

bigExcel = pd.read_excel("big.xlsx")
bigExcel['projectName'] = bigExcel["Ref."].apply(lambda x: "-".join(str(x).split("/")[:-1]))
groupByProjectName = bigExcel.groupby("projectName")

for name, group in groupByProjectName:
    group: pd.DataFrame = group
    group.rename(columns={"Ref.": "PO NO."}, inplace=True)
    del group["projectName"]
    group.sort_values(by='Doc. No.')
    outstanding = group["Outstanding"].sum()
    print(len(group))

    outdict = {"Outstanding": outstanding}
    group = group.append(outdict, ignore_index=True)
    group.to_excel(name+"-"+str(len(group))+".xlsx", index=None)
