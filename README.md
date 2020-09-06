# Excel to Json
excel转换成json的工具，在原仓库的基础上扩展解析规则  

### TODO List  
- [x] 读取Excel表  
- [x] 解析规则  
- [x] 控制台参数  

### 使用方式
将项目生成的_Console.exe复制到有excel文件的文件夹，会自动将文件夹内所有.xlsx文件转换为json文件并输出该目录下的"json"文件夹
```
@echo off
cd %~dp0
FabricioEx.exe -o json
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
解决方案：右键项目，选择“管理NuGet程序包”，或按照[官网]（https://www.nuget.org/packages/Newtonsoft.Json）指导安装
