
# -*- coding: utf-8 -*-

import os
import pandas as pd
import re
from io import StringIO


if __name__ == "__main__":
    
    folder_name = "Z:/jbian/PROJECT_DTNXP/KELU/thomas/testdata/ttest/"

    file_names = os.listdir(folder_name)


    writer = pd.ExcelWriter('Z:/jbian/PROJECT_DTNXP/KELU/thomas/testdata/40_GUI.xlsx')
    #data = pd.read_csv('F:/2022 work/Custom/kelu/thomas/testdata/HIOKI/40_DATA//IC1.csv', encoding='utf-8')
    for file_name in file_names:
        data = pd.read_csv('Z:/jbian/PROJECT_DTNXP/KELU/thomas/testdata/ttest/'+ file_name, encoding='utf-8')
        data.to_excel(writer, file_name, index=False)


    print('数据输出成功')
    writer.save()
    writer.close()

'''
# -*- coding: utf-8 -*-
import os
import pandas as pd
import re
from io import StringIO, BytesIO


if __name__ == "__main__":
    
    #base_path = "F:/2022 work/Custom/kelu/thomas/testdata/HIOKI/40_DATA/"
    base_path = "Z:/jbian/PROJECT_DTNXP/KELU/thomas/testdata/ttest/"
    out_file = base_path + '40_GUI.xlsx'

    file_names = os.listdir(base_path)

    writer = pd.ExcelWriter(out_file)
    #data = pd.read_csv('F:/2022 work/Custom/kelu/thomas/testdata/HIOKI/40_DATA//IC1.csv', encoding='utf-8')
    for file_name in file_names:
        u_path_name = base_path + file_name
        if (base_path + file_name) != out_file: #check input output file are same.
            if os.path.splitext(file_name)[-1] == ".csv":
                f = open(u_path_name, 'r')
                fd = f.read()
                #data = pd.read_csv(StringIO(fd), encoding="utf-8")
                data = pd.read_csv(u_path_name, encoding="utf-8")
                f.close()
            else:
                data = pd.read_excel(u_path_name)
            #print (type(data))
            #data.to_csv('e:/out.csv')
            data.to_excel(writer, os.path.splitext(file_name)[0], index=False)
            break

    print('数据输出成功')
    writer.save()
    writer.close()

'''