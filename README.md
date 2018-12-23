# LemonData
A solution to realize data configuration
一个数据配置和生成的解决方案

一句话概括：该方案用于解决游戏开发中的数据配置 及 客户端与服务器交互的协议 的自动化生成工作

1.与服务器交互，采用protobuf的方式，不清楚的可自行Google，在proto文件夹中，有定义的proto和批处理GenCS.bat

批处理内容如下：

```
echo on

set Path=ProtoGen\protogen.exe

for /f "delims=" %%m in ('dir /b "*.proto"') do %Path%  -i:%%m    -o:OutPut/proto.cs    -q  -d

pause
```

GenCS.bat会处理所在目录的proto文件，将其生成proto.cs文件，输出目录指定为同级OutPut/下


2.配置表分为定义文件和配置文件，均使用.xml格式。优点比单纯使用.excel作为配置文件或者.json文件作为配置文件，在可读性、文件大小等方面有较大的优势。
而客户端和服务端则使用.json文件，我提供了.py脚本用来将.xml生成.cs和json。

配置表文件结构见Template文件夹

\Template\define目录下有若干.xml格式的定义文件，内容大致如下：

```
<?xml version="1.0" encoding="utf-8"?>
<Avatar key = "id"> 
	<id type = "Int32" comment="头像id" />
	<scale type = "float" comment="缩放" />
	<name type = "string" comment="头像名称" />
	<url type = "string" comment="资源url" />
</Avatar>
```


除了.xml文件外，还有2个.py文件GenTemplate.py和GenTemplate_py2x.py

python 3.x使用GenTemplate.py

python 2.x使用GenTemplate_py2x.py

双击即可，会处理所在目录及其子目录下所有.xml文件，生成.cs和.xml

打开.py文件即可配置输出路径等信息

输出的.xml即为配置文件，大致如下
```
<?xml version="1.0" encoding="UTF-8"?>
<root xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
	<data>
		<Avatar>
			<entry>
				<REMARK0>#Avatar</REMARK0>
				<id>头像id</id>
				<scale>缩放</scale>
				<name>头像名称</name>
				<url>资源url</url>
				<REMARK1>-</REMARK1>
				<REMARK2>-</REMARK2>
			</entry>
			<entry>
				<REMARK0>#类型</REMARK0>
				<id>Int32</id>
				<scale>float</scale>
				<name>string</name>
				<url>string</url>
			</entry>
			<entry>
				<id>1</id>
				<scale>1</scale>
				<name>剑圣</name>
				<url>ui://Common</url>
			</entry>
		</Avatar>
	</data>
</root>
```
可以使用Excel或者Notepad++等文本工具编辑，然后运行同目录下的GenJson.py文件
会将json文件输出到指定目录，编辑GenJson.py即可配置输出路径
输出内容如下：
```
{"entry":[
{"id":1,"scale":1.5,"name":"剑圣","url":"ui://Common/jiansheng"},
{"id":2,"scale":1.5,"name":"武器大师","url":"ui://Common/wuqidashi"},
{"id":3,"scale":1.5,"name":"兰博","url":"ui://Common/lanbo"},
{"id":4,"scale":1.5,"name":"薇恩","url":"ui://Common/weien"}
]}
```
