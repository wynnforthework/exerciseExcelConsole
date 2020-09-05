# Excel to Json
excel转换成json的工具，在原仓库的基础上扩展解析规则

### 使用方式
将项目生成的_Console.exe复制到有excel文件的文件夹，会自动将文件夹内所有.xlsx文件转换为json文件并输出该目录下的"json"文件夹
```
@echo off
cd %~dp0
_Console.exe -o json
pause
```

### 目录结构
xlsx2json:.  
│  Guide.xlsx  
│  _Console.exe  
│  _转换并拷贝到项目.bat  
│   
└─json  
           Guide.json  
        
### 本地调试错误及解决方法
1. 编译时软件报错：error CS0234: 命名空间“Microsoft”中不存在类型或命名空间名“  
解决方案：点击项目->添加引用->程序集->扩展  
选中软件提示缺少的组件，我选的是Microsoft.Office.Interop.Excel，点击确定  