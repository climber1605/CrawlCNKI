#!/usr/bin/env python3
import sys
import os
import shutil
from pathlib import Path
import pandas as pd
import collections

# 从src(excel文件)中读取期刊名称，读取结果保存到dst中
def excel2txt(src, dst):
    if Path(dst).is_file():
        return
    data_frame = pd.read_excel(src, sheet_name='待下载列表', skiprows=1, usecols=[1])
    journals = list(data_frame.to_dict(orient='dict')['期刊名称'].values())
    # 将读出的期刊名称转存为txt文件
    with open(dst, 'w') as f: 
        for journal in journals:  
            f.write(journal)  
            f.write('\n')

# 从filename文件中读取出第start到end行
def read_txt(filename, start, end):
    line_num = 0
    result = []
    with open(filename, 'r', encoding='UTF-8') as f:
        line = f.readline()
        while line != '':
            if start <= line_num <= end:
                result.append(line[:-1])
            line_num += 1
            line = f.readline()
    return result
         
# 找出src目录下的文件相比target多出与缺失的文件，将多出的文件移动到dst目录下 
def find_extra_missing(src, target, dst):
    if not os.path.exists(dst):
        os.mkdir(dst)

    src_files = set(os.listdir(src))
    target_files = set([journal + str(i + 1) + '.xlsx' for journal in target for i in range(3)])
    moved_files, missing_files = [], []
    for src_file in src_files:
        if src_file not in target_files:
            shutil.move(src + '/' + src_ｆile, dst)
            moved_files.append(src_file)
            print('extra file: {}, moved: {}, total: {}'.format(src_file, len(moved_files), len(src_files)))
    for target_file in target_files:
        if target_file not in src_files:
            missing_files.append(target_file)
            print('missing file: {}, missing: {}, total: {}'.format(target_file, len(missing_files), len(target_files)))
    return moved_files, missing_files

# 将src目录下的文件按期刊名合并，合并后的文件保存到dst目录下
def merge_journals(src, dst):
    print('Start merging files in {}'.format(src))
    if not os.path.exists(dst):
        os.mkdir(dst)

    file2df = collections.defaultdict(list)
    num_before = num_after = 0
    # 遍历文件目录，将所有表格表示为pandas中的DataFrame对象
    for root_dir, sub_dir, files in os.walk(src):
        for file in files:
            num_before += 1
            if file.endswith('xlsx'):
            	# 构造绝对路径
                file_name = os.path.join(root_dir, file)
                df = pd.read_excel(file_name)
                journal_file = file[:-6] + '.xlsx'
                file2df[journal_file].append(df)
    num_after = len(file2df)

    for file in file2df:
        df_concated = pd.concat(file2df[file])
        #df_concated.drop_duplicates(subset=['篇名'], keep='first', inplace=True)
        out_path = os.path.join(dst, file)
        df_concated = df_concated.loc[:, ~df.columns.str.contains('Unnamed')]
        df_concated = df_concated.sort_values(by='发表时间')
        df_concated.index = range(1, len(df_concated) + 1)
        df_concated.to_excel(out_path, sheet_name='Sheet1', index_label='序号')

    print('Finish merging files in {}. There are {} files before merge, {} files after merge.'.format(src, num_before, num_after))    

# 将src目录下的所有文件合并，合并后的文件保存到dst目录下
def merge_publish_numbers(src, dst):
    print('Start merging files in {}'.format(src))
    if not os.path.exists(dst):
        os.mkdir(dst)
    
    res = []
    num_before = 0
    # 遍历文件目录，将所有表格表示为pandas中的DataFrame对象
    for root_dir, sub_dir, files in os.walk(src):
        for file in files:
            num_before += 1
            if file.endswith('xlsx'):
            	# 构造绝对路径
                file_name = os.path.join(root_dir, file)
                df = pd.read_excel(file_name)
                res.append(df)

    df_concated = pd.concat(res)
    df_concated.drop_duplicates(subset=['期刊名称'], keep='first', inplace=True)
    df_concated = df_concated.loc[:, ~df.columns.str.contains('Unnamed')]
    output_file = os.path.join(dst, '发表数量.xlsx')
    df_concated.index = range(1, len(df_concated) + 1)
    df_concated.to_excel(output_file, sheet_name='Sheet1', index_label='序号')
    print('Finish merging files in {}. There are {} files before merge.'.format(src, num_before))  

# 检查每个期刊2012年-2020年总文献数量是否符合预期，返回不符合预期的期刊名列表
def check_publish_numbers(src, dst):
    print('Start checking files in {}'.format(src))
    invalid = []
    total = 0
    for root_dir, sub_dir, files in os.walk(src):
        for file in files:
            total += 1
            src_file = os.path.join(root_dir, file)
            df_src = pd.read_excel(src_file)
            dst_file = os.path.join(dst, file)
            df_dst = pd.read_excel(dst_file)
            if len(df_src) != df_dst['总数'][0]:
                print('Expected number: {}, actual number: {}, file: {}'.format(df_dst['总数'][0], len(df_src), file))
                invalid.append(file)
    print('Finish checking files in {}. Valid: {}, invalid: {}, total: {}'.format(src, total - len(invalid), len(invalid), total))
    return invalid

if __name__ == '__main__':
    if (len(sys.argv)) != 3:
        print('Usage: python3 utils.py start end')
        exit(0)
    else:
        start, end = int(sys.argv[1]), int(sys.argv[2])

    src, dst = './output' + '_' +str(start) + '_' + str(end), './merged_output' + '_' + str(start) + '_' + str(end)
    excel2txt('./待爬取数据.xlsx', './journals.txt')
    target = read_txt('./journals.txt', start, end)

    moved_files, missing_files = find_extra_missing(src, target, './others')
    print('extra files: {}'.format(moved_files))
    print('missing files: {}'.format(missing_files))
    
    merge_journals(src, dst)
    pub_dir = './publish_numbers'
    print(check_publish_numbers(dst, pub_dir))
    #merge_publish_numbers(pub_dir, './')
    
