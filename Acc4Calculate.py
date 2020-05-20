import numpy as np
import pandas as pd
import os

def readDataSet(field=True):
    '''
    :param field: 是否为现场数据, True-现场数据, False-云平台数据
    :return: super_short_term - 超短期数据, short_term - 短期数据
    '''
    cur_path = os.getcwd()
    name = 'Name Error'
    super_short_term = pd.DataFrame()
    short_term = pd.DataFrame()
    available_power = pd.DataFrame()
    if field == True:
        dataset_path = cur_path + '\现场数据\超短期单点数据集'
        files = os.listdir(dataset_path)
        if not files:
            print('There is no Dataset in ...\现场数据\超短期单点数据集, please check it!')
            return
        for file in files:
            if '超短期单点预测' in file:
                super_short_term_single_points = pd.read_excel(dataset_path + '\\' + file)
                super_short_term = pd.concat([super_short_term, super_short_term_single_points.iloc[:, [0, 2, 3, 4, 5, 6, 7, 8, 9, 1]]], axis=0)
        dataset_path = cur_path + '\现场数据'
        files = os.listdir(dataset_path)
        super_short_term = super_short_term.reset_index(drop=True)
        for file in files:
            if '理论功率15分钟' in file:
                name = file[24:26] + file[27:29] + '.xlsx'
                theo_power = pd.read_excel(dataset_path + '\\' + file)
                available_power = theo_power.iloc[:, 3]
                super_short_term = pd.concat([super_short_term, available_power], axis=1)
        for file in files:
            if '短期功率统计' == file[:6]:
                power_stat = pd.read_excel(dataset_path + '\\' + file, header=1)
                short_term = pd.concat([short_term, power_stat.iloc[:, 0], power_stat.iloc[:, 3], power_stat.iloc[:, 1]], axis=1)
        short_term = pd.concat([short_term, available_power], axis=1)
    else:
        dataset_path = cur_path + '\云平台数据'
        files = os.listdir(dataset_path)
        for file in files:
            if '超短期单点预测' in file:
                super_short_term_single_points = pd.read_excel(dataset_path + '\\' + file)
                name = file[23:25] + file[26:28] + '.xlsx'
                super_short_term = pd.concat([super_short_term, \
                                    super_short_term_single_points.iloc[:, [0, 2, 3, 4, 5, 6, 7, 8, 9]], \
                                    super_short_term_single_points.iloc[:, 1]], axis=1)
        for file in files:
            if '理论功率15分钟' in file:
                theo_power = pd.read_excel(dataset_path + '\\' + file)
                available_power = theo_power.iloc[:, 3]
                super_short_term = pd.concat([super_short_term, available_power], axis=1)
        for file in files:
            if '风速功率对比统计' in file:
                speed_VS_power = pd.read_excel(dataset_path + '\\' + file)
                short_term = pd.concat([short_term, speed_VS_power.iloc[:, 0], speed_VS_power.iloc[:, 2], speed_VS_power.iloc[:, 1]], axis=1)
        short_term = pd.concat([short_term, available_power], axis=1)

    return super_short_term, short_term

def calc_and_2Excel(super_st, st, cap, base_coef=0.03, name='tmp.xlsx'):
    '''
    :param super_st: super shot term 超短期数据
    :param st: short term 短期数据
    :param cap: 装机容量，MW
    :param base_coef: 预测值跟可用值都大于cap*base_coef的情况下才计入统计, default to 0.03
    :param name: file name of the final .xlsx, default to tmp.xlsx
    :return: super_st, st (with calculated accuracy)
    :return: df_acc, a dataframe like this:
    :return:        |    '15分钟精度', '30分钟精度', '45分钟精度', '60分钟精度'  |
    :return:        |        acc1        acc2          acc3         acc4     |
    :return: acc5 - mean accuracy for |预测-可用|/可用 of short-term-data
    :return: acc6 - mean accuracy for |预测-实发|/实发 of short-term-data
    :return: acc7 - mean accuracy for |可用-实发|/实发 of short-term-data
    '''
    base = cap * base_coef

    # 处理超短期数据
    super_st['预测1-可用'] = super_st['15分钟'] - super_st['机头风速可用功率']
    super_st.loc[(super_st['15分钟'] < base) & (super_st['机头风速可用功率'] < base), '|预测1-可用|/可用'] = float('inf')
    super_st.loc[((super_st['15分钟'] >= base) | (super_st['机头风速可用功率'] >= base)) \
                 & (super_st['机头风速可用功率'] == 0), '|预测1-可用|/可用'] = 1
    super_st.loc[((super_st['15分钟'] >= base) | (super_st['机头风速可用功率'] >= base)) \
                 & (super_st['机头风速可用功率'] != 0), '|预测1-可用|/可用'] = \
                 np.abs(super_st['预测1-可用']) / super_st['机头风速可用功率']
    super_st.loc[(super_st['|预测1-可用|/可用'] != float('inf')) & (super_st['|预测1-可用|/可用'] > 1), '|预测1-可用|/可用'] = 1
    super_st['预测1-实发'] = super_st['15分钟'] - super_st['实发功率']


    super_st['预测2-可用'] = super_st['30分钟'] - super_st['机头风速可用功率']
    super_st.loc[(super_st['30分钟'] < base) & (super_st['机头风速可用功率'] < base), '|预测2-可用|/可用'] = float('inf')
    super_st.loc[((super_st['30分钟'] >= base) | (super_st['机头风速可用功率'] >= base)) \
                 & (super_st['机头风速可用功率'] == 0), '|预测2-可用|/可用'] = 1
    super_st.loc[((super_st['30分钟'] >= base) | (super_st['机头风速可用功率'] >= base)) \
                 & (super_st['机头风速可用功率'] != 0), '|预测2-可用|/可用'] = \
                 np.abs(super_st['预测2-可用']) / super_st['机头风速可用功率']
    super_st.loc[(super_st['|预测2-可用|/可用'] != float('inf')) & (super_st['|预测2-可用|/可用'] > 1), '|预测2-可用|/可用'] = 1
    super_st['预测2-实发'] = super_st['30分钟'] - super_st['实发功率']

    super_st['预测3-可用'] = super_st['45分钟'] - super_st['机头风速可用功率']
    super_st.loc[(super_st['45分钟'] < base) & (super_st['机头风速可用功率'] < base), '|预测3-可用|/可用'] = float('inf')
    super_st.loc[((super_st['45分钟'] >= base) | (super_st['机头风速可用功率'] >= base)) \
                 & (super_st['机头风速可用功率'] == 0), '|预测3-可用|/可用'] = 1
    super_st.loc[((super_st['45分钟'] >= base) | (super_st['机头风速可用功率'] >= base)) \
                 & (super_st['机头风速可用功率'] != 0), '|预测3-可用|/可用'] = \
                 np.abs(super_st['预测3-可用']) / super_st['机头风速可用功率']
    super_st.loc[(super_st['|预测3-可用|/可用'] != float('inf')) & (super_st['|预测3-可用|/可用'] > 1), '|预测3-可用|/可用'] = 1
    super_st['预测3-实发'] = super_st['45分钟'] - super_st['实发功率']

    super_st['预测4-可用'] = super_st['60分钟'] - super_st['机头风速可用功率']
    super_st.loc[(super_st['60分钟'] < base) & (super_st['机头风速可用功率'] < base), '|预测4-可用|/可用'] = float('inf')
    super_st.loc[((super_st['60分钟'] >= base) | (super_st['机头风速可用功率'] >= base)) \
                 & (super_st['机头风速可用功率'] == 0), '|预测4-可用|/可用'] = 1
    super_st.loc[((super_st['60分钟'] >= base) | (super_st['机头风速可用功率'] >= base)) \
                 & (super_st['机头风速可用功率'] != 0), '|预测4-可用|/可用'] = \
                 np.abs(super_st['预测4-可用']) / super_st['机头风速可用功率']
    super_st.loc[(super_st['|预测4-可用|/可用'] != float('inf')) & (super_st['|预测4-可用|/可用'] > 1), '|预测4-可用|/可用'] = 1
    super_st['预测4-实发'] = super_st['60分钟'] - super_st['实发功率']

    super_st['可用-实发'] = super_st['机头风速可用功率'] - super_st['实发功率']

    acc1_all = np.array(super_st['|预测1-可用|/可用'])
    acc1_all = acc1_all[acc1_all != float('inf')]
    acc1 = 1 - np.mean(acc1_all)

    acc2_all = np.array(super_st['|预测2-可用|/可用'])
    acc2_all = acc2_all[acc2_all != float('inf')]
    acc2 = 1 - np.mean(acc2_all)

    acc3_all = np.array(super_st['|预测3-可用|/可用'])
    acc3_all = acc3_all[acc3_all != float('inf')]
    acc3 = 1 - np.mean(acc3_all)

    acc4_all = np.array(super_st['|预测4-可用|/可用'])
    acc4_all = acc4_all[acc4_all != float('inf')]
    acc4 = 1 - np.mean(acc4_all)

    df_acc = pd.DataFrame(data=np.array([acc1, acc2, acc3, acc4]), columns=['精度'])
    df_acc.index = pd.Series(['15分钟', '30分钟', '45分钟', '60分钟'])

    super_st.loc[super_st['|预测1-可用|/可用'] == float('inf'), '|预测1-可用|/可用'] = ''
    super_st.loc[super_st['|预测2-可用|/可用'] == float('inf'), '|预测2-可用|/可用'] = ''
    super_st.loc[super_st['|预测3-可用|/可用'] == float('inf'), '|预测3-可用|/可用'] = ''
    super_st.loc[super_st['|预测4-可用|/可用'] == float('inf'), '|预测4-可用|/可用'] = ''

    # 处理短期数据
    st['预测-可用'] = st['预测功率'] - st['机头风速可用功率']
    st.loc[(st['预测功率'] < base) & (st['机头风速可用功率'] < base), '|预测-可用|/可用'] = float('inf')
    st.loc[((st['预测功率'] >= base) | (st['机头风速可用功率'] >= base)) \
                 & (st['机头风速可用功率'] == 0), '|预测-可用|/可用'] = 1
    st.loc[((st['预测功率'] >= base) | (st['机头风速可用功率'] >= base)) \
                 & (st['机头风速可用功率'] != 0), '|预测-可用|/可用'] = \
        np.abs(st['预测-可用']) / st['机头风速可用功率']
    st.loc[(st['|预测-可用|/可用'] != float('inf')) & (st['|预测-可用|/可用'] > 1), '|预测-可用|/可用'] = 1

    st['预测-实发'] = st['预测功率'] - st['实发功率']
    st.loc[(st['预测功率'] < base) & (st['实发功率'] < base), '|预测-实发|/实发'] = float('inf')
    st.loc[((st['预测功率'] >= base) | (st['实发功率'] >= base)) \
           & (st['实发功率'] == 0), '|预测-实发|/实发'] = 1
    st.loc[((st['预测功率'] >= base) | (st['实发功率'] >= base)) \
           & (st['实发功率'] != 0), '|预测-实发|/实发'] = \
        np.abs(st['预测-实发']) / st['实发功率']
    st.loc[(st['|预测-实发|/实发'] != float('inf')) & (st['|预测-实发|/实发'] > 1), '|预测-实发|/实发'] = 1

    st['可用-实发'] = st['机头风速可用功率'] - st['实发功率']
    st.loc[(st['机头风速可用功率'] < base) & (st['实发功率'] < base), '|可用-实发|/实发'] = float('inf')
    st.loc[((st['机头风速可用功率'] >= base) | (st['实发功率'] >= base)) \
           & (st['实发功率'] == 0), '|可用-实发|/实发'] = 1
    st.loc[((st['机头风速可用功率'] >= base) | (st['实发功率'] >= base)) \
           & (st['实发功率'] != 0), '|可用-实发|/实发'] = \
        np.abs(st['可用-实发']) / st['实发功率']
    st.loc[(st['|可用-实发|/实发'] != float('inf')) & (st['|可用-实发|/实发'] > 1), '|可用-实发|/实发'] = 1

    acc5_all = np.array(st['|预测-可用|/可用'])
    acc5_all = acc5_all[acc5_all != float('inf')]
    acc5 = 1 - np.mean(acc5_all)

    acc6_all = np.array(st['|预测-实发|/实发'])
    acc6_all = acc6_all[acc6_all != float('inf')]
    acc6 = 1 - np.mean(acc6_all)

    acc7_all = np.array(st['|可用-实发|/实发'])
    acc7_all = acc7_all[acc7_all != float('inf')]
    acc7 = 1 - np.mean(acc7_all)

    st.loc[st['|预测-可用|/可用'] == float('inf'), '|预测-可用|/可用'] = ''
    st.loc[st['|预测-实发|/实发'] == float('inf'), '|预测-实发|/实发'] = ''
    st.loc[st['|可用-实发|/实发'] == float('inf'), '|可用-实发|/实发'] = ''

    # write to Excel
    tail = [['' for i in range(24)] for j in range(2)]
    tail[1][12] = acc1
    tail[1][15] = acc2
    tail[1][18] = acc3
    tail[1][21] = acc4
    super_st_tail = pd.DataFrame(tail, columns=super_st.columns)
    super_st = pd.concat([super_st, super_st_tail])

    tail = [['' for i in range(10)] for j in range(2)]
    tail[1][5] = acc5
    tail[1][7] = acc6
    tail[1][9] = acc7
    st_tail = pd.DataFrame(tail, columns=st.columns)
    st = pd.concat([st, st_tail])

    writer = pd.ExcelWriter(name)
    super_st.to_excel(writer, sheet_name='超短期', index=False)
    st.to_excel(writer, sheet_name='短期', index=False)
    df_acc.to_excel(writer, sheet_name='sheet3', index=True)
    writer.save()

    return super_st, st, df_acc, (acc1, acc2, acc3, acc4, acc5, acc6, acc7)




def write2Excel():
    pass

if __name__ == "__main__":
    super_st, st = readDataSet(filed=True)
    calc_and_2Excel(super_st, st, cap=402, name='现场.xlsx')