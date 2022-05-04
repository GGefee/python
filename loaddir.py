def movefile(file1):
    import os
    import shutil
    #file_sr = '1.xlsx'#调试用
    file_sr = file1#发布用
    dirname, filename = os.path.split(file_sr)
    #文件所在目录
    file_full_path = os.path.dirname(os.path.abspath(__file__))
    print (">>>>> 2. The file folder path: os.path.dirname(os.path.abspath(__file__)) path is:", file_full_path)
    data = filename.split(".")
    #print len(data) 
    print (data[0])
    print (data[1])
    #获取当前所有文件和文件夹
    #os.listdir(file_full_path)
        #print('已存在')
        #print(i,j,k)
    #else:
    #    os.mkdir(str(data[0]))



    if not os.path.exists(data[0]):
            # 目录不存在，进行创建操作
            os.makedirs(data[0]) #使用os.makedirs()方法创建多层目录
            print("目录新建成功：" + data[0])
    else:
            print("目录已存在！！！")
    return data[0]




    '''
        if str(data[0]) in dirs:
            print('已存在该目录')
        else:
            print('不存在该目录')
            os.mkdir(data[0])
        
    # 获得文件夹及文件完整路径
    for root, dirs, files in os.walk(file_full_path,topdown=False):
        print(files)
    '''