#!/usr/bin/env python
# coding: utf-8

# In[1]:


# pip install openpyxl
import pandas as pd


# book = pd.read_excel('./../file/test.xlsx')
filename= 'C:/國家教育研究院學生學習成就資料釋出資料/國家教育研究院學生學習成就資料釋出資料/臺灣學生學習成就評量資料/2009/數學科選擇題試題作答反應/小四/2009_小四_數學科_原始資料檔.xls'
xl = pd.ExcelFile(filename)


# In[2]:


sheet_list = xl.sheet_names #讀取資料清單
sheet_list


# In[3]:


data_dict = pd.read_excel(filename, sheet_name=sheet_list)
data_dict


# In[4]:


combine_data = pd.DataFrame()
combine_data = pd.concat(data_dict.values(), ignore_index=True)
combine_data


# In[5]:


combine_data.describe()


# In[6]:


combine_data.info()


# In[7]:



# 資料列的值 =9 變成空值(NaN)

filter_cols = combine_data.columns[2:len(combine_data.columns)+1]
print(filter_cols)


# In[8]:


# 塞選我們要的欄位

combine_data[filter_cols]


# In[9]:


combine_data[filter_cols]

# 挑選符合條件True的資料列，不符合條件False
combine_data[filter_cols].isin([1,2,3,4])

#不符合條件False的變為Nan
stu_data = combine_data[combine_data[filter_cols].isin([1,2,3,4])]

print(stu_data)


# In[10]:


stu_data['學生代號'] = combine_data['學生代號'] 
stu_data['題本號'] = combine_data['題本號'] 

stu_data[['學生代號','題本號']] = combine_data[['學生代號','題本號']]
stu_data


# In[11]:


#移除空值
stu_data = stu_data.dropna()
stu_data


# In[12]:


combine_data.info()


# In[14]:


import pandas as pd


# book = pd.read_excel('./../file/test.xlsx')
reference_table =pd.read_excel('C:/國家教育研究院學生學習成就資料釋出資料/國家教育研究院學生學習成就資料釋出資料/臺灣學生學習成就評量資料/2009/數學科選擇題試題作答反應/小四/2009_小四_數學科_試題區塊對照表(改).xls',sheet_name='04試題區塊對照表')


# In[15]:


reference_table.head()


# # D 資料表(學生原始資料&試題區塊對照表)合併

# In[16]:


for i in range(1,len(stu_data.columns)-1):
    origin_name = '試題{}'.format(i)
    rename = '第{}題'.format(i)
    stu_data = stu_data.rename(columns={origin_name: rename})


# In[17]:


sstu_data = stu_data.melt(id_vars=['學生代號','題本號'])

stu_data


# In[18]:


# test 一行轉換重塑

o_reference_table = pd.DataFrame()
i = 1
test_col_name = '題本0{}'.format(i)
reference_table = reference_table.rename(columns={test_col_name:i})
o_reference_table = reference_table.melt(value_vars=[i])
print(o_reference_table)


# In[19]:


#test

ans_col_name = '正確答案'
o_reference_table['正解'] = reference_table[ans_col_name]
o_reference_table['題號'] = reference_table['題號']
print(o_reference_table)


# In[20]:


#test
i = 2
test_col_name ='題本0{}'.format (i)
reference_table = reference_table.rename (columns={test_col_name: i})
t_reference_table = reference_table.melt(value_vars=[i])
print(t_reference_table)


ans_col_name ='正確答案.{}'.format(i-1)
t_reference_table['正解'] = reference_table [ans_col_name]
t_reterence_tab1e['題號'] = reference_table['題號']
print (t_reference_table)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              


# 2.資料表格格式轉換
# 2.1寬資料表->長資料表 (stu_data -> long_stu_data)
# melt() (參數： id_vars：指定不要進行重塑的欄位; value_vars：指定哪些欄位要重塑成單一欄位)

# In[21]:


stu_data = stu_data.melt(id_vars=['學生代號','題本號'])

# 資料表轉換後,修改資料欄位名稱
stu_data = stu_data.rename(columns={'variable': '題號','value':'學生作答'})

stu_data


# In[22]:


reference_table.head()


# In[23]:


o_reference_table = pd.DataFrame()
for i in range(1,len(sheet_list)+1):
    t_reference_table = pd.DataFrame()
    if i < 10:
        test_col_name ='題本0{}'.format(i)
        # 題本01 --> 1 , 題本02 --> 2
        reference_table = reference_table.rename(columns={test_col_name: i})
        
        if i == 1:
            o_reference_table =reference_table.melt(value_vars=[i])
            ans_col_name ='正確答案'
            o_reference_table['正解'] = reference_table[ans_col_name]
            o_reference_table['題號'] = reference_table['題號']
        else:
            t_reference_table = reference_table.melt(value_vars=[i])
            
            ans_col_name ='正確答案.{}'.format(i-1)
            t_reference_table['正解'] = reference_table[ans_col_name]
            t_reference_table['題號'] = reference_table['題號']
            # 疊合資料表
            o_reference_table = pd.concat([o_reference_table,t_reference_table], ignore_index=True)
    else:
        test_col_name ='題本{}'.format(i)
        reference_table = reference_table.rename(columns={test_col_name: i})
        
        t_reference_table = reference_table.melt(value_vars=[i])
        
        ans_col_name ='正確答案.{}'.format(i-1)
        t_reference_table['正解'] = reference_table[ans_col_name]
        t_reference_table['題號'] = reference_table['題號']
        
        o_reference_table = pd.concat([o_reference_table,t_reference_table], ignore_index=True)


# In[24]:


o_reference_table


# In[25]:


full_refer_table = o_reference_table.rename(columns={'variable':'題本號','value':'試題區塊'})


# In[26]:


full_refer_table.head()


# In[27]:


new_df = pd.merge(stu_data, full_refer_table,  how='left')


# In[28]:


new_df


# In[29]:


#E.判斷學生作答是否正確
result_list = []

for i in range(0,len(new_df)):
    if (new_df['學生作答'][i] == new_df['正解'][i]):
        result_list.append(1)
    else:
        result_list.append(0)


# In[30]:


new_df['ma結果'] = result_list

print(new_df)

new_df.describe()


# 2009_小四_國語文_試題指標與區塊對照(整).xlsx
# 1. test_group_table存入

# In[34]:


test_group_table = pd.read_excel("C:/國家教育研究院學生學習成就資料釋出資料/國家教育研究院學生學習成就資料釋出資料/臺灣學生學習成就評量資料/2009/數學科選擇題試題作答反應/小四/2009_小四_數學科_試題指標與區塊對照(整).xlsx")
stu_c_data = pd.merge(new_df, test_group_table,  on='試題區塊')


# In[35]:


test_group_table


# In[36]:


# G.計算學生題本答對率-1


stu_test = stu_c_data.groupby(['學生代號','題本號'],as_index=False)['ma結果'].sum()

stu_test


# In[37]:


# G.計算學生題本答對率-2. score_rate_list存入題本答對率結果

#題號數量
test_length = len(reference_table['題號'])

#計算題本答對率
score_rate_list = []
for i in range(0,len(stu_test)):
    score_rate = round((stu_test['ma結果'][i]/test_length)*100,2)
    score_rate_list.append(score_rate)

# 新增欄位
stu_test['ma題本答對率'] = score_rate_list


# In[38]:


stu_test


# In[39]:


# H.學生在每個指標的答對率-1


stu_class_group = stu_c_data.groupby(['學生代號','題本號','大指標'],as_index=False)['ma結果'].sum()


# In[40]:


stu_class_group


# In[41]:


# H.學生在每個指標的答對率-2 計算題本的指標題數
test_class_group = stu_c_data.groupby(['學生代號','題本號','大指標'],as_index=False)['題號'].count()


# In[42]:


test_class_group


# In[43]:


# H.學生在每個指標的答對率-3 資料合併


# class_group_df = pd.merge(stu_class_group, test_class_group,  how='left')

class_group_df = pd.merge(stu_class_group, test_class_group,  on=['學生代號','題本號','大指標'])


# In[44]:


class_group_df


# In[45]:


# H.學生在每個指標的答對率-4. 計算學生每個指標的答對率


#指標清單

class_no = class_group_df['大指標'].unique()

print(class_no)


# In[46]:


# H.學生在每個指標的答對率-2 計算題本的指標題數
test_class_group = stu_c_data.groupby(['學生代號','題本號','大指標'],as_index=False)['題號'].count()

test_class_group = test_class_group.rename(columns={'題號': '指標題數'})


# In[47]:


test_class_group


# In[48]:


test_class_group = test_class_group.rename(columns={'題號': '指標題數'})


# In[49]:


class_group_df = pd.merge(stu_class_group, test_class_group,  on=['學生代號','題本號','大指標'])


# In[50]:


no = 1

print(class_group_df['大指標']== no)

groups = class_group_df.loc[class_group_df['大指標'] == no]
print(groups)


# In[51]:


class_no = class_group_df['大指標'].unique()
print(class_no)


# In[52]:


#計算指標答對率
for no in class_no:
    class_score_rate_list = []
    groups = class_group_df.loc[class_group_df['大指標'] == no]
    
    for g in range(0, len(groups)):
        stu_id = groups.iloc[g]['學生代號']
        score = groups.iloc[g]['ma結果']
        q_num = groups.iloc[g]['指標題數']
       
        class_score_rate = [stu_id  ,  round((score /q_num)*100,2)]
        class_score_rate_list.append(class_score_rate)

    column = 'ma指標{}答對率'.format(no)
    df = pd.DataFrame(class_score_rate_list, columns =['學生代號', column])
    stu_test = pd.merge(stu_test, df,  how='left')


# In[53]:


stu_test


# In[55]:


# I.匯出csv

#### UTF-8-Sig和UTF-8的主要差別是前者是UTF-8 with BOM (Byte Order Mark)，在Win10中翻譯成「具有BOM的UTF-8」，後者沒有BOM，總之有BOM的比較好。

stu_test.to_csv('C:/國家教育研究院學生學習成就資料釋出資料/國家教育研究院學生學習成就資料釋出資料/臺灣學生學習成就評量資料/2009/數學科選擇題試題作答反應/小四/data.csv',encoding='utf-8-sig'  )


# In[ ]:




