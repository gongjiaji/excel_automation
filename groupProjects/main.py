import pandas as pd
import rich
import time, datetime
from styleframe import StyleFrame, Styler, utils
import os

# remove historical results first
cur_path1 = "Result/"
ls = os.listdir(cur_path1)
for i in ls:
    c_path = os.path.join(cur_path1, i)
    os.remove(c_path)

bigExcel = pd.read_excel("big.xlsx", dtype={"Doc. No.": str})
bigExcel["projectName"] = bigExcel["Ref."].apply(
    lambda x: "-".join(str(x).split("/")[:-1])
)
groupByProjectName = bigExcel.groupby("projectName")

for name, group in groupByProjectName:
    group: pd.DataFrame = group.sort_values(by="Doc. No.")
    group.rename(columns={"Ref.": "PO NO.", "Doc. No.": "Inv No."}, inplace=True)
    del group["projectName"]
    outstanding = group["Outstanding"].sum()
    outdict = {"Outstanding": outstanding}
    group = group.append(outdict, ignore_index=True)

    # styler
    len_max = {x: group.get(x).astype(str).str.len().max() * 1.25 for x in group}
    default_style = Styler(
        horizontal_alignment=utils.horizontal_alignments.left,
        shrink_to_fit=False,
        wrap_text=False,
    )
    sf = StyleFrame(group, styler_obj=default_style)
    sf.apply_column_style(
        cols_to_style="Date",
        styler_obj=Styler(number_format=utils.number_formats.date),
    )
    ew = StyleFrame.ExcelWriter("Result/" + name + "-" + str(len(group) - 1) + ".xlsx")
    StyleFrame.A_FACTOR = 0
    sf.to_excel(excel_writer=ew, best_fit=[x for x in group])
    ew.save()
    print("finish", name)
print("DONE!")
