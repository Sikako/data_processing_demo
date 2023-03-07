#!/usr/bin/env python
# coding: utf-8

# In[445]: 記得安裝套件
import pandas as pd
import os, xlsxwriter


# In[446]: 各區段的衛星名稱
Index1 = ['GPSELE', 'GPSAZI', 'GPSS1C', 'GPSS2S', 'GPSS2W', 'GPSS5Q']
Index2 = ['GALELE', 'GALAZI', 'GALS1C', 'GALS5Q', 'GALS7Q', 'GALS8Q']
Index3 = ['GLOELE', 'GLOAZI', 'GLOS1C', 'GLOS2C', 'GLOS2P']
Index4 = ['BDSELE', 'BDSAZI', 'BDSS2I', 'BDSS5P', 'BDSS7I']
current_row = 0


# 檔案讀寫區，只要更改inputname、outputname
inputname = '0104.xlsx'
outputname = '0104_output.xlsx'
# ---------------------------------------------------------
if not os.path.exists(outputname):
    workbook = xlsxwriter.Workbook(outputname)
    workbook.close()
writer = pd.ExcelWriter(outputname, engine="xlsxwriter")


# In[447]: 分類器，初始給值Classficator(上面的Index陣列, 每顆衛星數量x#, excel表格從第幾列開始, excel表格從第幾列結束)
class Classficator():
    # 初始函式
    def __init__(self, Index, satellite_num, start_range, stop_range):
        self.Index = Index
        self.satellite_num = satellite_num
        self.start_range = start_range
        self.stop_range = stop_range
        self.step = len(Index)
        self.x_list = list(map(lambda x: "x" + str(x).zfill(2), range(1, satellite_num+1)))
        self.df_total_sorted = None
    
    # 讀檔功能
    def read_excel(self, excel_name, sheet_name):
        df_total = pd.DataFrame(columns=self.Index)
        df_raw = pd.read_excel(open(excel_name, "rb"), sheet_name=sheet_name)
        df_raw = df_raw.replace('-', 0)
        df_raw.index += 2
        for i in range(self.start_range,self.stop_range+1, self.step):
            df_section = df_raw[i-2:i-2+self.step][self.x_list]
            df_section_nonempty = df_section.loc[:, (df_section != 0).any(axis=0)]
            df_section_nonempty = df_section_nonempty.transpose()
            # print(df_section_nonempty, "\n")
            df_section_nonempty.columns = self.Index
            df_total =  pd.concat([df_total, df_section_nonempty], ignore_index=True)
        self.df_total_sorted = df_total.sort_values(by=self.Index[0])
    
    # 獲取分類dataFrame-------------------------------------------------------------------
    def get_i_to_i_plus_n_df(self, i, n):
        df = self.df_total_sorted[i <= self.df_total_sorted[self.Index[0]]]
        df = df[df[self.Index[0]] <= i+n-1].transpose()
        df = df.rename(columns={x:y for x,y in zip(df.columns, range(1, len(df.columns)+1))})
        df.index.name = f"{i}-{i+n-1}"
        return df
    # ---------------------------------------------------------------------------------------
    
    # 寫檔功能
    def write_excel(self, StartSectionNum, StopSectionNum, n):
        global current_row
        for i in range(StartSectionNum, StopSectionNum+1, n):
            self.get_i_to_i_plus_n_df(i, n).to_excel(writer, startrow=current_row)
            current_row += (self.step+2)

        # self.get_0_30_df().to_excel(writer, startrow=current_row)
        # current_row += (self.step+2)
        # self.get_31_50_df().to_excel(writer, startrow=current_row)
        # current_row += (self.step+2)
        # self.get_51_70_df().to_excel(writer, startrow=current_row)
        # current_row += (self.step+2)
        # self.get_71_90_df().to_excel(writer, startrow=current_row)
        # current_row += (self.step+2)


# In[448]:


classficator1 = Classficator(Index1, 45, 2, 98)
classficator1.read_excel(inputname, "同時間比較")
classficator1.write_excel(1, 81, 10)


# In[449]:


classficator2 = Classficator(Index2, 45, 104, 200)
classficator2.read_excel(inputname, "同時間比較")
classficator2.write_excel(1, 81, 10)


# In[450]:


classficator3 = Classficator(Index3, 45, 206, 286)
classficator3.read_excel(inputname, "同時間比較")
classficator3.write_excel(1, 81, 10)


# In[451]:


classficator4 = Classficator(Index4, 45, 291, 371)
classficator4.read_excel(inputname, "同時間比較")
classficator4.write_excel(1, 81, 10)


# In[452]:


writer.save()

