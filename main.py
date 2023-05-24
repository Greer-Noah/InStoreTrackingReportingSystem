import customtkinter
import InterfaceCreation
import os
import pandas as pd
import xlsxwriter

interface = InterfaceCreation.InterfaceCreation(customtkinter.CTk, 800, 650)
interface.updateInterface()
print("STORE LIST: ", interface.store_list)

report_file_name = "WeeklyReport{}.xlsx".format(interface.store_list[-1].get_date_input())

path = os.path.join(os.path.expanduser("~"),
                    "Downloads/TrackingReports_{}".format(interface.store_list[-1].get_date_input()))
if not os.path.exists(path):
    os.mkdir(path)

path = os.path.join(os.path.expanduser("~"),
                    "Downloads/TrackingReports_{}".format(interface.store_list[-1].get_date_input()),
                    report_file_name)



if os.path.exists(path):
    filename, extension = os.path.splitext(path)
    counter = 1
    while os.path.exists(path):
        path = filename + " (" + str(counter) + ")" + extension
        counter += 1

str(path)

global writer
writer = pd.ExcelWriter(path, engine='xlsxwriter')

combined_matching_sheet_name = "Combined Matching"
interface.store_list[-1].get_combined().to_excel(writer, combined_matching_sheet_name, startrow=0, startcol=0,
                                                 index=False)
combined_qb_matching_sheet_name = "Combined QB Matching"
interface.store_list[-1].get_qb_combined().to_excel(writer, combined_qb_matching_sheet_name, startrow=0, startcol=0,
                                                 index=False)
combined_repl_sheet_name = "Combined REPL Breakdown"
interface.store_list[-1].get_combined_repl().to_excel(writer, combined_repl_sheet_name, startrow=0, startcol=0,
                                                 index=False)

for store in interface.store_list:
    print("Gathering Store {}'s Data...".format(store.store_num))
    matching_sheet_name = "Matching {}".format(store.get_store_num())
    str(matching_sheet_name)
    qb_matching_sheet_name = "QB Matching {}".format(store.get_store_num())
    str(qb_matching_sheet_name)
    total_items_sheet_name = "Total Items {}".format(store.get_store_num())
    str(total_items_sheet_name)
    expected_items_sheet_name = "Expected Items {}".format(store.get_store_num())
    str(expected_items_sheet_name)
    repl_group_nbr_sheet_name = "REPL Breakdown {}".format(store.get_store_num())
    str(repl_group_nbr_sheet_name)

    store.get_matching().to_excel(writer, matching_sheet_name, startrow=0, startcol=0, index=False)
    store.get_qb_matching().to_excel(writer, qb_matching_sheet_name, startrow=0, startcol=0, index=False)
    store.get_total_items().to_excel(writer, total_items_sheet_name, startrow=0, startcol=0, index=False)
    store.get_expected().to_excel(writer, expected_items_sheet_name, startrow=0, startcol=0, index=False)
    store.get_repl_nbr().to_excel(writer, repl_group_nbr_sheet_name, startrow=0, startcol=0, index=False)

writer.save()

count = 0
for store in interface.store_list:
    print("Exporting Store {} Cycle Count Output...".format(interface.store_list[count].get_store_num()))
    cycle_count_workbook_name = "Store{}CycleCountOutput.xlsx".format(interface.store_list[count].get_store_num())
    output_file_path = os.path.join(os.path.expanduser("~"),
                        "Downloads/TrackingReports_{}".format(interface.store_list[-1].get_date_input()),
                                    cycle_count_workbook_name)
    writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')
    workbook = writer.book
    worksheet1 = workbook.add_worksheet('Unique EPCs, Dupe UPCs')
    worksheet2 = workbook.add_worksheet('Unique UPCs')
    number_format = workbook.add_format({'num_format': '0'})
    worksheet1.set_column('B:B', None, number_format)
    worksheet2.set_column('A:A', None, number_format)

    store.get_UE().to_excel(writer, "Unique EPCs, Dupe UPCs", startrow=0, startcol=0, index=False)
    store.get_DU().to_excel(writer, "Unique EPCs, Dupe UPCs", startrow=0, startcol=1, index=False)
    store.get_UU().to_excel(writer, "Unique UPCs", startrow=0, startcol=0, index=False)
    store.get_error_EPCs().to_excel(writer, "Errors", startrow=0, startcol=0, index=False)
    store.get_error_messages().to_excel(writer, "Errors", startrow=0, startcol=1, index=False)

    writer.save()
    count += 1

print("All files exported. Exiting...")
raise SystemExit(0)