# ReadMe.md
### 功能详情

1. 这是一个对已有excel末尾数据进行追加的函数

### 使用说明

1. 参数含义
   1. excel_name : 源文件
   2. new_excel_name ：生成的新文件,若不需要可以在源码中修改下save函数的保存文件路径变量
   3. insert_data   ：插入的数据 数据类型：dict类型  键为Sheet的名字，值为需要插入的参数列表。
      1. eg：    insert_data = {"Sheet1" : [["第1行数据列表"],["第2行数据列表"]....]} 以此类推