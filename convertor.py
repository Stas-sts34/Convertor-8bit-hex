import openpyxl
import pandas as pd
import os


class Converter:

    def convertion(self, number):
        revers_number = ''.join(reversed(number))
        # revers_number = f"{(14 - len(revers_number)) * '0'}{revers_number}"     #add null, if len str <14
        while len(revers_number) < 14:
            revers_number = revers_number + "0"
        binar_list = []
        for i in revers_number:
            bin_str = bin(int(i))[2:]
            bin_str = f"{(3 - len(bin_str)) * '0'}{bin_str}"
            # while len(str(bin_str)) < 3:
            #     bin_str = '0' + bin_str
            binar_list.append(bin_str)
        out_binar = ''.join(binar_list)
        out_hex = str(hex(int(out_binar, 2)))
        out_hex = out_hex[-6:]
        return out_hex.upper()

    def add_null(self, in_number):
        in_number = f"{(14 - len(in_number)) * '0'}{in_number}"     #add null, if len str <14
        return in_number

    def open_xls(self, file, file_out):
        data_xls = pd.read_excel(file,
                                 sheet_name='Лист1',
                                 index_col=0)      # конвертирут из xlsx в csv
        data_xls.to_csv('in.csv', encoding='cp1251')
        df_convert = pd.read_csv('in.csv', encoding='cp1251', dtype=str)
        df = pd.DataFrame(columns=['0', '1'])          # создание нового датафрейма
        for i in range(0, len(df_convert)):
            number_in = self.add_null(df_convert.iloc[i, 0])
            number_out = self.convertion(str(df_convert.iloc[i, 0]))
            df.loc[i] = number_in, number_out       # заполнение датафрейма
        os.remove('in.csv')
        try:
            df.to_excel(file_out,  index=False)
        except PermissionError:
            df.to_excel(file_out + '1', index=False)


# elem = Converter()
# elem.open_xls('in.xlsx', 'out3.xlsx')
