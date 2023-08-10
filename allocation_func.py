import pandas as pd
import numpy as np


pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.width', 5000)
pd.options.mode.chained_assignment = None
def all_func(name_alloc_):
    input_bussines = pd.read_excel(name_alloc_)
    input_bussines = input_bussines.loc[~input_bussines['Функционал-юнит (источник)'].isin(['нет в'])].reset_index(drop=True)

    aggr_data = input_bussines.groupby(by=['Тип связи', 'Функционал-юнит (источник)', 'Наименование функционал-юнита (источник)', 'Наименование бизнеса',
                                        'Код бизнеса'], as_index=False)[['факт', 'план']].sum()
    input_bussines = aggr_data


    article = pd.read_excel('тип передачи_2.xlsx')
    input_bussines['план'] = input_bussines['план'].fillna(0)
    input_bussines['факт'] = input_bussines['факт'].fillna(0)
    input_bussines['виконання %'] = 0
    for i in range(len(input_bussines)):
        if input_bussines['план'][i] != 0:
            input_bussines['виконання %'][i] = round(input_bussines['факт'][i] / input_bussines['план'][i] * 100)
        else:
            input_bussines['виконання %'][i] = 100
    input_bussines['план'] = input_bussines['план'].replace(1, 0)
    input_bussines['факт'] = input_bussines['факт'].fillna(0)
    input_bussines['вдіхилення'] = input_bussines['факт'] - input_bussines['план']

    table_names = input_bussines['Наименование бизнеса'].drop_duplicates()
    table_names = list(sorted(table_names))

    def transposition(value, num_sv):
        new_df = input_bussines[input_bussines['Тип связи'] == num_sv]
        wide_titanic = new_df.pivot(
            index=['Функционал-юнит (источник)', 'Наименование функционал-юнита (источник)'],

            columns='Наименование бизнеса',
            values=value
        )
        wide_titanic.reset_index(drop=False,  inplace=True)
        return wide_titanic

    lst_num = ["план", 'факт', "вдіхилення", "виконання %",]
    lst_sv = ['1.2.', '1.7.', '1.1.']

    concat_df_1_2 = pd.concat([transposition(lst_num[i], lst_sv[0]) for i in range(len(lst_num))], axis=1, ignore_index=True)
    article_1_2 = article[article['Тип передачи'] == '1.2.']
    df_1_2_ = pd.merge(left=article_1_2, right=concat_df_1_2, left_on='источник', right_on=concat_df_1_2.iloc[:, 0], how='left')


    concat_df_1_7 = pd.concat([transposition(lst_num[i], lst_sv[1]) for i in range(len(lst_num))], axis=1, ignore_index=True)
    article_1_7 = article[article['Тип передачи'] == '1.7.']
    df_1_7_ = pd.merge(left=article_1_7, right=concat_df_1_7, left_on='источник', right_on=concat_df_1_7.iloc[:, 0], how='left')

    concat_df_1_1 = pd.concat([transposition(lst_num[i], lst_sv[2]) for i in range(len(lst_num))], axis=1, ignore_index=True)
    article_1_1 = article[article['Тип передачи'] == '1.1.']
    df_1_1_ = pd.merge(left=article_1_1, right=concat_df_1_1, left_on='источник', right_on=concat_df_1_1.iloc[:, 0], how='left')

    concat_dfs = pd.concat([df_1_2_, df_1_7_, df_1_1_], axis=0)



    concat_dfs['источник'] = concat_dfs['источник'].fillna(concat_dfs.iloc[:, 0])
    concat_df = concat_dfs.drop(concat_dfs.columns[[0, 3, 4]], axis=1)


    def moving(x):
        data = concat_df.iloc[:, [x, x+7, x+14, x+21]]
        return data
    lst_count_buss = [2,3,4,5,6]

    concat_df_ = pd.concat([moving(lst_count_buss[i]) for i in range(len(lst_count_buss))], axis=1, ignore_index=True)
    concat_df_2 = pd.concat([concat_df.iloc[:, :2], concat_df_], axis=1)


    lst_columns = ['Код статьи', 'Наименование статьи'] + lst_num * len(table_names)
    concat_df_2.columns = lst_columns

    concat_df_2.fillna(0, inplace=True)

    similar = concat_df_2.reset_index()


    def summa(num, i):
        while similar['index'][i] != 0:
            if i+1 != len(similar):
                i += 1
            else:
                break
        similar.iloc[num, 3:] = similar.iloc[num:i+1, 3:].sum(axis=0)
        return similar
    lst_con = [0, 23, 34]
    axe = 0
    for i in range(len(lst_con)):
        axe = summa(lst_con[i], lst_con[i]+1)
    axe = axe.drop(axe.columns[[0]], axis=1)

    def func(tini_df):
        tini_df['план_'] = tini_df.iloc[:, 2::4].sum(axis=1)
        tini_df['факт_'] = tini_df.iloc[:, 3::4].sum(axis=1)
        tini_df['вдіхилення_'] = tini_df.iloc[:, 4::4].sum(axis=1)
        tini_df['виконання %_'] = tini_df.iloc[:, 5::4].sum(axis=1)

        tini_df['_план_'] = tini_df.iloc[:, [2,10]].sum(axis=1)
        tini_df['_факт_'] = tini_df.iloc[:, [3, 11]].sum(axis=1)
        tini_df['_вдіхилення_'] = tini_df.iloc[:, [4,12]].sum(axis=1)
        tini_df['_виконання %_'] = tini_df.iloc[:, [5, 13]].sum(axis=1)


        consol_bank = tini_df.iloc[:, 22:26]
        corp_buss = tini_df.iloc[:, 26:30]
        corp_net = tini_df.iloc[:, 10:14]
        vip = tini_df.iloc[:, 2:6]
        msb = tini_df.iloc[:, 14:18]
        rozdrib = tini_df.iloc[:, 18:22]
        inshi = tini_df.iloc[:, 6:10]

        conclusion = pd.concat([tini_df.iloc[:, [0,1]], consol_bank, corp_buss, corp_net, vip, msb, rozdrib, inshi], axis=1)
        return conclusion


    conclusion = func(axe)
    conclusion.iloc[:,[2, 3, 4, 5]] = 0


    conclusion.loc[-1] = {'Код статьи': '!7', 'Наименование статьи' : 'Аллокації'}
    conclusion.index = conclusion.index + 1
    conclusion = conclusion.sort_index()


    name_code = conclusion[['Наименование статьи']]
    name_state = conclusion[['Код статьи']]
    conclusion = conclusion.set_index('Код статьи')
    conclusion = conclusion.drop(['Наименование статьи'], axis=1)
    conclusion.loc['!7'] = conclusion.loc['1.2.'] + conclusion.loc['1.7.'] + conclusion.loc['1.1.']

    conclusion.reset_index(
                drop=False,  # False означает, что существующий индекс переместится в колонки, а не будет удален
                inplace=True  # произвести операцию на месте, без присваивания новой переменной
            )

    res_finish_all = pd.concat([name_state, name_code, conclusion.iloc[:, 1:]], axis=1)
    res_finish_all.to_excel('res_all_.xlsx')
    return res_finish_all

