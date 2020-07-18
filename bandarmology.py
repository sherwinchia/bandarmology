import pandas as pd
import numpy as np


def newDataFrame(link, header_number, column_array):
    new_data_frame = pd.read_excel(
        link, header=header_number, usecols=column_array)
    return new_data_frame


def reformatExcel(data_frame, buy_column_name, sell_column_name):
    data_frame = data_frame.astype('str')
    data_frame = data_frame.apply(lambda x: x.str.replace(',', ''))
    data_frame[buy_column_name] = data_frame[buy_column_name].astype('float')
    data_frame[sell_column_name] = data_frame[sell_column_name].astype('float')
    data_frame.sort_values(by=buy_column_name, ascending=False)
    return data_frame


def calculateNetVolumeLot(data_frame, column_name):
    total_lot = data_frame[column_name].sum()
    return str(total_lot)


def calculateTop(total_broker, data_frame, buy_column_name, sell_column_name):
    status = np.where(data_frame[buy_column_name].head(total_broker).sum(
    ) > -(data_frame[sell_column_name].head(total_broker).sum()), 'ACCUMULATION', 'DISTRIBUTION')
    return str(status)


def calculateNetLot(total_broker, data_frame, buy_column_name, sell_column_name):
    total_lot = data_frame[buy_column_name].head(total_broker).sum(
    )+data_frame[sell_column_name].head(total_broker).sum()
    return total_lot


def calculateNetRatio(net_lot, net_volume_lot):
    net_ratio = net_lot / float(net_volume_lot) * 100
    return net_ratio


def displayNetRatioStatus(net_ratio):
    if net_ratio > 0:
        if net_ratio > 30:
            return "BIG ACCUMULATION"
        elif net_ratio <= 30 and net_ratio > 10:
            return "ACCUMULATION"
        elif net_ratio <= 10:
            return "SMALL ACCUMULATION"
    else:
        if net_ratio >= -10:
            return "SMALL DISTRIBUTION"
        elif net_ratio >= -30 and net_ratio < -10:
            return "DISTRIBUTION"
        elif net_ratio <= -30:
            return "BIG DISTRIBUTION"


# Baca dan reformat file excel
data_frame = newDataFrame('sample1.xls', 2, [0, 1, 2, 3, 5, 6, 7, 8])

# nama column yang kita mau (sementara maybe(?))
buy_column_name = list(data_frame.columns)[1]
sell_column_name = list(data_frame.columns)[5]

# Reformat column dan ubah datatype ke float
data_frame = reformatExcel(data_frame, buy_column_name, sell_column_name)

# Dipakai untuk calculate net ratio
net_volume_lot = calculateNetVolumeLot(data_frame, buy_column_name)

# Display tabel
print(data_frame)

# Display net volume untuk lot
print("\nNET VOLUME(LOT): " + str(net_volume_lot) + "\n")

# Display net lot per top x
print("NET (LOT):")
for x in range(5):
    top_status = calculateTop(
        x+1, data_frame, buy_column_name, sell_column_name)
    net_lot = calculateNetLot(
        x+1, data_frame, buy_column_name, sell_column_name)
    print("TOP " + str(x+1) + ": " + str(top_status) +
          " | Net Lot: " + str(net_lot))

# Display net ratio per top x
print("\nNET RATIO:")
for x in range(5):
    net_ratio = calculateNetRatio(calculateNetLot(
        x+1, data_frame, buy_column_name, sell_column_name), net_volume_lot)
    print("TOP " + str(x+1) + ": " + displayNetRatioStatus(net_ratio) +
          " | Net Ratio: " + str(net_ratio))