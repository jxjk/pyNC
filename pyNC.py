import pandas as pd  
from datetime import datetime 
  
# 读取Excel文件并生成NC宏程序 无条件类型 
def generate_nc_macro_from_excel_1(file_path,o_path,macro):  
    # 读取Excel文件，header=0表示第一行是表头，作为DataFrame的列名  
    df = pd.read_excel(file_path, index_col=0,header=0, engine='openpyxl') 
    # 提取注释行（假设第二行是注释）  
    try:  
        comments = df.iloc[0].astype(str).tolist()  # 将注释转换为字符串列表  
    except IndexError:  
        # 如果没有第二行，可以设置一个默认注释或跳过此步骤  
        comments = []   
    # 忽略第二行的注释（如果有的话）  
    data = df.iloc[1:]  
    # 获取当前日期并格式化为字符串  
    current_date = datetime.now().strftime('%Y.%m.%d')  
    # 初始化NC宏程序字符串  
    nc_macro = []  
    # 添加程序头部，包含当前日期  
    nc_macro.append('%')  
    nc_macro.append(f'O{o_path}({file_path})')  
    nc_macro.append(f'(DATE {current_date})')  
    nc_macro.append('/M1')  
    nc_macro.append('(*********)')  
    nc_macro.append(f'GOTO#{macro}') 
    nc_macro.append('#3000=1(***DATA ERR***)')  
    # 遍历数据行生成NC宏的每一部分  
    for index, row in data.iterrows():  
        n_line_number = index 
        nc_macro.append(f'N{n_line_number}')   
        i=0
        for col in df.columns:  # 遍历列名（宏变量名）  
            value = row[col]  # 获取当前行的列值  
            nc_macro.append(f'{col}={value}{comments[i]}')  # 生成赋值语句  
            i+=1
        nc_macro.append('GOTO99') 
    nc_macro.append('N99\n')    
    nc_macro.append('M99\n')    
    nc_macro.append('%\n')    
    # 将列表转换为字符串，每个元素占一行  
    nc_macro_str = '\n'.join(nc_macro)  
    # 输出或保存到文件  
    with open(f'./out/O{o_path}.NC', 'w') as f:  
        f.write(nc_macro_str)  
    return nc_macro_str  


# 读取Excel文件并生成NC宏程序 条件类型 
def generate_nc_macro_from_excel_2(file_path,o_path):  
    # 读取Excel文件，header=0表示第一行是表头，作为DataFrame的列名  
    df = pd.read_excel(file_path, index_col=0,header=0, engine='openpyxl') 
    # 提取注释行（假设第二行是注释）  
    try:  
        comments = df.iloc[0].astype(str).tolist()  # 将注释转换为字符串列表  
    except IndexError:  
        # 如果没有第二行，可以设置一个默认注释或跳过此步骤  
        comments = []   
    # 忽略第二行的注释（如果有的话）  
    data = df.iloc[1:]  
    # 获取当前日期并格式化为字符串  
    current_date = datetime.now().strftime('%Y.%m.%d')  
    # 初始化NC宏程序字符串  
    nc_macro = []  
    # 添加程序头部，包含当前日期  
    nc_macro.append('%')  
    nc_macro.append(f'O{o_path}({file_path})')  
    nc_macro.append(f'(DATE {current_date})')  
    nc_macro.append('/M1')  
    nc_macro.append('(*********)')  
    # nc_macro.append('GOTO#503') 
    i=1
    for index, row in data.iterrows():  
        nc_macro.append(f'{index}{i*10}')
        i +=1
    nc_macro.append('#3000=1(***DATA ERR***)')  
    # 遍历数据行生成NC宏的每一部分  
    j=1
    for index, row in data.iterrows():  
        n_line_number = j*10 
        nc_macro.append(f'N{n_line_number}')   
        j+=1
        k=0
        for col in df.columns:  # 遍历列名（宏变量名）  
            value = row[col]  # 获取当前行的列值  
            nc_macro.append(f'{col}={value}{comments[k]}')  # 生成赋值语句  
            k+=1
        nc_macro.append('GOTO99') 
    nc_macro.append('N99\n')    
    nc_macro.append('M99\n')    
    nc_macro.append('%\n')    
    # 将列表转换为字符串，每个元素占一行  
    nc_macro_str = '\n'.join(nc_macro)  
    # 输出或保存到文件  
    with open(f'./out/O{o_path}.NC', 'w') as f:  
        f.write(nc_macro_str)  
    return nc_macro_str  

 
# 读取Excel文件并生成NC宏程序 垫高块 
def generate_nc_macro_from_excel_3(file_path,o_path):  
    # 读取Excel文件，header=0表示第一行是表头，作为DataFrame的列名  
    df = pd.read_excel(file_path, index_col=0,header=0, engine='openpyxl') 
    # 提取注释行（假设第二行是注释）  
    try:  
        comments = df.iloc[0].astype(str).tolist()  # 将注释转换为字符串列表  
    except IndexError:  
        # 如果没有第二行，可以设置一个默认注释或跳过此步骤  
        comments = []   
    # 忽略第二行的注释（如果有的话）  
    data = df.iloc[1:]  
    # 获取当前日期并格式化为字符串  
    current_date = datetime.now().strftime('%Y.%m.%d')  
    # 初始化NC宏程序字符串  
    nc_macro = []  
    # 添加程序头部，包含当前日期  
    nc_macro.append('%')  
    nc_macro.append(f'O{o_path}({file_path})')  
    nc_macro.append(f'(DATE {current_date})')  
    nc_macro.append('(#800:dianGaoKuai_chunChuZhi)')  
    nc_macro.append('/M1')  
    nc_macro.append('(*********)')  
    # nc_macro.append('GOTO#503') 
    i=1
    for index, row in data.iterrows():  
        nc_macro.append(f'{index}{i*10}')
        i +=1
    nc_macro.append('#3000=1(***DATA ERR***)')  
    # 遍历数据行生成NC宏的每一部分  
    j=1
    for index, row in data.iterrows():  
        n_line_number = j*10 
        nc_macro.append(f'N{n_line_number}')   
        j+=1
        # k=0
        for col in df.columns:  # 遍历列名（宏变量名）  
            value = row[col]  # 获取当前行的列值  
            nc_macro.append(f'{col}={value}') # {comments[k]}')  # 生成赋值语句  
            # k+=1

        nc_macro.append('IF[#800EQ#508]GOTO99')  
        nc_macro.append(f'#3006=#508(CHANGE PLANT)')  
        nc_macro.append('#800=#508')  
        nc_macro.append('GOTO99')  
        
    nc_macro.append('\nN99')    
    nc_macro.append('M99')    
    nc_macro.append('%\n')    
    # 将列表转换为字符串，每个元素占一行  
    nc_macro_str = '\n'.join(nc_macro)  
    # 输出或保存到文件  
    with open(f'./out/O{o_path}.NC', 'w') as f:  
        f.write(nc_macro_str)  
    return nc_macro_str 

def toolInfo():
    nc_macro_str = '''
        %
        O1900(TOOL INFO)
        (DATE 2016.7.5)
        (TOOL INFO)
        #1=171(TOOL NO. STRAT)
        #2=185(TOOL NO. END)
        (LOOP)
        WH[#1LE#2]DO1
            #152=#[#1](SET TOOL H)
            M98P1300(COUNDITION)

            IF[#[800+#151]EQ#152]]GOTO9(如果刀具正在使用就跳过)

            T#151(换刀)
            M6
            #3006=#152(CHANGE TOOL)
            #[800+#151]=#152(设置刀具记忆)
            N9
            #1=#1+1
        END1

        N99
        M99
        %\n
        '''
    with open(f'./out/O1900.NC', 'w') as f:  
        f.write(nc_macro_str)  
    return nc_macro_str 


# 读取Excel文件并生成NC宏程序 加工条件材料分类选择、主程序选择 
def generate_nc_macro_from_excel_4(file_path,o_path,macro=515):  
    # 读取Excel文件，header=0表示第一行是表头，作为DataFrame的列名  
    df = pd.read_excel(file_path, index_col=0,header=0, engine='openpyxl') 
    # 提取注释行（假设第二行是注释）  
    try:  
        comments = df.iloc[0].astype(str).tolist()  # 将注释转换为字符串列表  
    except IndexError:  
        # 如果没有第二行，可以设置一个默认注释或跳过此步骤  
        comments = []   
    # 忽略第二行的注释（如果有的话）  
    data = df.iloc[1:]  
    # 获取当前日期并格式化为字符串  
    current_date = datetime.now().strftime('%Y.%m.%d')  
    # 初始化NC宏程序字符串  
    nc_macro = []  
    # 添加程序头部，包含当前日期  
    nc_macro.append('%')  
    nc_macro.append(f'O{o_path}({file_path})')  
    nc_macro.append(f'(DATE {current_date})')  
    nc_macro.append(f'(#{macro})')  
    nc_macro.append('/M1')  
    nc_macro.append('(*********)')  
    nc_macro.append(f'GOTO#{macro}') 
    nc_macro.append('#3000=1(***DATA ERR***)')  
    # 遍历数据行生成NC宏的每一部分  

    for index, row in data.iterrows():  
        n_line_number = index 
        nc_macro.append(f'N{n_line_number}')   
        i=0
        for col in df.columns:  # 遍历列名（宏变量名）  
            value = row[col]  # 获取当前行的列值  
            nc_macro.append(f'M98P{value}')  # 生成赋值语句  
            i+=1
        nc_macro.append('GOTO99') 
    nc_macro.append('\nN99')    
    if o_path == 1100:
        nc_macro.append('M30')    
    else:
        nc_macro.append('M99')    
    nc_macro.append('%\n')    
    # 将列表转换为字符串，每个元素占一行  
    nc_macro_str = '\n'.join(nc_macro)  
    # 输出或保存到文件  
    with open(f'./out/O{o_path}.NC', 'w') as f:  
        f.write(nc_macro_str)  
    return nc_macro_str  


# 主程序入口  
def main():  
    # Excel文件路径  
    # 生成NC宏程序  
    excel_file_path7 = 'Type_input_0.xlsx'  
    nc_macro = generate_nc_macro_from_excel_4(excel_file_path7,1100,502)  # 依据类型参数，选择主程序。

    excel_file_path0 = 'caiLiao_input.xlsx'  
    nc_macro += generate_nc_macro_from_excel_1(excel_file_path0,1009,502)  # 依据类型参数，选择材料。 

    excel_file_path1 = 'type_input_1.xlsx'  
    nc_macro += generate_nc_macro_from_excel_1(excel_file_path1,1101,503)  # 依据类型参数，输出关联属性1。

    excel_file_path2 = 'type_input_2.xlsx'  
    nc_macro += generate_nc_macro_from_excel_2(excel_file_path2,1102)  # 依据类型参数，输出关联属性2。

    excel_file_path3 = 'tool_input.xlsx'  
    nc_macro+= generate_nc_macro_from_excel_2(excel_file_path3,1201)  # 依据类型参数，选择刀具。

    nc_macro += toolInfo()                                             #  O1900 输出刀具提示程序 

    excel_file_path5 = 'dianGaoKuai_input.xlsx'  
    nc_macro += generate_nc_macro_from_excel_3(excel_file_path5,1401)  # 依据类型参数，选择垫块，并提示更换。

    excel_file_path6 = 'tiaoJianType_input.xlsx'  
    nc_macro += generate_nc_macro_from_excel_4(excel_file_path6,1300,515)  # 依据材质别，选择材料别加工条件。

    excel_file_path4 = 'tiaoJian_input_1.xlsx'  
    nc_macro += generate_nc_macro_from_excel_1(excel_file_path4,1301,152)  # 材质1 依据刀具H号，选择加工条件。

    excel_file_path8 = 'tiaoJian_input_2.xlsx'  
    nc_macro += generate_nc_macro_from_excel_1(excel_file_path8,1302,152)  # 材质2 依据刀具H号，选择加工条件。

    excel_file_path9 = 'tiaoJian_input_3.xlsx'  
    nc_macro += generate_nc_macro_from_excel_1(excel_file_path9,1303,152)  # 材质3 依据刀具H号，选择加工条件。

    excel_file_path10 = 'tiaoJian_input_4.xlsx'  
    nc_macro += generate_nc_macro_from_excel_1(excel_file_path10,1304,152)  # 材质4 依据刀具H号，选择加工条件。


    with open(f'./out/All.NC', 'w') as f:  
        f.write(nc_macro)  

  
if __name__ == '__main__':  
    main()