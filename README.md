# Excel to Json
excel转换成json的工具，在原仓库的基础上扩展解析规则  

### TODO List  
- [x] 读取Excel表  
- [x] 解析规则  
- [x] 控制台参数  

### 使用方式
创建.bat批处理文件，将以下代码复制到文件中保存到exe相同的目录中，将项目生成的FabricioEx.exe复制到有excel文件的文件夹，会自动将文件夹内所有.xlsx文件转换为json文件并输出该目录下的"json"文件夹  
新增：  
-h 帮助文档  
-b 是否允许单元格为空（第一行第一列不能为空）   
-o 输出目录自动判断是否为觉得路径，如果不是，则以当前目录为根目录的相对路径，没有-o参数，则默认输出到当前目录。  
-w 是否监控当前目录下文件的改变（窗口不关闭的情况下，文件改变后会自动转换成json）
```
@echo off
cd %~dp0
FabricioEx.exe -o json -b n -w n
pause
```

### 表格规范（解析规则） 
<table>
<thead><tr><th>Excel</th><th>Json</th><th>用处</th></tr></thead><tbody>
 <tr>
 <td>ID</td>
 <td rowspan="2">{"id": "0"}</td>
 <td rowspan="2">number,string等基本类型的简单数据结构</td>
 </tr>
 <tr><td>1</td></tr>
 <tr><td>Params@</td><td rowspan="2">{"Params": ["1","2","3"]}</td><td rowspan="2">集合数据</td></tr>
 <tr><td>1,2,3</td></tr>
 <tr><td>Options@#anchor</td><td rowspan="2">{"Options": [{"Override_SceneName_FunctionName_Judge":["SelectRoomScene","onSelect","evt.room==3"]},{"Equal":["this.hero_id==3"]}</td><td rowspan="2">集合对象</td></tr>
 <tr><td>1</td></tr>
 <tr><td>Labels#anchor</td><td rowspan="2">{"Labels": {"title":"势如破竹","content":"三级战斗房间，英雄技能伤害+1%"}}</td><td rowspan="2">单个对象</td></tr>
 <tr><td>1</td></tr>
 <tr><td>!Desc</td><td rowspan="2">该列会被忽略</td><td rowspan="2">注释</td></tr>
 <tr><td>首页</td></tr>
</tbody></table>

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

2. 增加json库  
解决方案：右键项目，选择“管理NuGet程序包”，或按照[官网](https://www.nuget.org/packages/Newtonsoft.Json)指导安装

3. 打包成独立的exe文件  
解决方案：默认生成的exe文件和dll文件是分开的，如果只复制exe文件到别的文件夹执行会提示找不到json库，需要通过NuGet库管理工具另外安装“Costura.Fody”库，安装完成之后，再生成exe文件就是包含dll的了，可以随意复制使用。
