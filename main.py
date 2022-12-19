import pandas as pd
import datetime as dt

DATE = dt.datetime(2022,10,31)
DATE_NULL = dt.datetime(1900,1,1)

#copy names of input files to open
PATH1 = 'C:/Users/tatya/Desktop/pyth_work/part1/Normerica Orders Scheduled v Shipped 16dec2022 700 simon.xlsx'
PATH2 = 'C:/Users/tatya/Desktop/pyth_work/part2/dock_schedule_16dec2022_700 simon.xlsx'

data_orders = pd.read_excel(PATH1)
data_scheduled = pd.read_excel(PATH2)

#clean dataset by booked yes by null and bad dates
clean_dataset = data_orders[(data_orders['booked'] == "Y")]
clean_dataset.loc[clean_dataset['actual_ship'] == '(null)', 'actual_ship'] = DATE_NULL
clean_dataset.loc[clean_dataset['sched_ship'] == '(null)', 'sched_ship'] = DATE_NULL
checking_dataset = clean_dataset[((clean_dataset['actual_ship'] < dt.datetime(2022,1,1)) & (clean_dataset['actual_ship'] != DATE_NULL))
                                |((clean_dataset['sched_ship'] < dt.datetime(2022,1,1)) & (clean_dataset['sched_ship'] != DATE_NULL))]

if len(checking_dataset.index) > 0:
    for i in range(len(checking_dataset.index)):
        print(clean_dataset.loc[checking_dataset.index[i]])
        clean_dataset.loc[checking_dataset.index[i], 'sched_ship'] = dt.datetime(2022,11,10)
        print(clean_dataset.loc[checking_dataset.index[i]])

clean_dataset=clean_dataset[((clean_dataset['actual_ship'] >= DATE) & (clean_dataset['sched_ship'] >= DATE))
                            |((clean_dataset['actual_ship'] >= DATE) & (clean_dataset['sched_ship'] == DATE_NULL))
                            |((clean_dataset['sched_ship'] >= DATE) & (clean_dataset['actual_ship'] == DATE_NULL))
                            |((clean_dataset['sched_ship'] == DATE_NULL) & (clean_dataset['actual_ship'] == DATE_NULL))]

clean_dataset.loc[clean_dataset['actual_ship'] == DATE_NULL, 'actual_ship'] = '(null)'
clean_dataset.loc[clean_dataset['sched_ship'] == DATE_NULL, 'sched_ship'] = '(null)'
print("Number of clean rows is ",len(clean_dataset))
clean_dataset.to_excel('C:/Users/tatya/Desktop/pyth_work/part1/merged_pivots.xlsx', sheet_name='clean_data')

scheduled_pivot = clean_dataset.pivot_table(index="warehouse", columns='sched_ship', values='ship_count', aggfunc='sum', margins=True)
scheduled_pivot_renamed = scheduled_pivot.rename(index={'BD1':'BD1_Sched','DB1':'DB1_Sched','DY1':'DY1_Sched',
                                                        'LE1':'LE1_Sched', 'PO1':'PO1_Sched','PX1':'PX1_Sched',
                                                        'All':'Total_Sched'})
actual_pivot = clean_dataset.pivot_table(index="warehouse", columns='actual_ship', values='ship_count', aggfunc='sum', margins=True)
actual_pivot_renamed = actual_pivot.rename(index={'BD1':'BD1_Shipped','DB1':'DB1_Shipped','DY1':'DY1_Shipped',
                                                  'LE1':'LE1_Shipped', 'PO1':'PO1_Shipped','PX1':'PX1_Shipped',
                                                  'All':'Total_Shipped'})

#Merge two pivots
scheduled_pivot_renamed = scheduled_pivot_renamed.transpose()
actual_pivot_renamed = actual_pivot_renamed.transpose()
all_data = pd.merge(scheduled_pivot_renamed, actual_pivot_renamed, how='outer', left_index=True,right_index=True)
all_data = all_data.transpose()
all_data.fillna(0, inplace=True)
all_data.sort_index(inplace=True)
all_data.to_excel('C:/Users/tatya/Desktop/pyth_work/part1/merged_pivots.xlsx', sheet_name='report')

#filter_c_p_to_r_BD1
filtered_dataset=data_scheduled[(data_scheduled['Status'] != "C")]
filtered_dataset.loc[filtered_dataset['Status'] == 'P', 'Status'] = 'R'

filtered_dataset.fillna('(blank)', inplace=True)
filtered_pivot = filtered_dataset.pivot_table(index=['Status', 'Freight Terms','Tender Status'], columns='SSD', values='Sales Order#', aggfunc='count', margins=True)
print("Number of clean rows is ", len(filtered_dataset))

#order need to be closed (for report 2) to excel
orders_need_to_book = []
entered_status = data_orders[(data_orders['booked'] == "N")]
report_table = entered_status.pivot_table(index=None, columns='sched_ship', values='ship_count', aggfunc='sum').round(0).to_dict()

total_need_order = 0
for i in range(len(list(report_table.keys()))):
    total_need_order += (list(report_table.values())[i]).get('ship_count')
    orders_need_to_book.append(str(list(report_table.keys())[i]) + " = " + str((list(report_table.values())[i]).get('ship_count')) + '\n')
orders_need_to_book.append("Grand Total = " + str(total_need_order))
orders_df = pd.DataFrame(orders_need_to_book)

with pd.ExcelWriter("C:/Users/tatya/Desktop/pyth_work/part2/pivot_status.xlsx") as writer:
    filtered_pivot.to_excel(writer, sheet_name="report_pivot", index=True)
    orders_df.to_excel(writer, sheet_name="need_to_book_manualy", index=False)
