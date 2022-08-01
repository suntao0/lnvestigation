# lnvestigation
抓取企查查数据

1.功能介绍
基于selenium对企查查信息和相关招标进行爬取数据，爬取后进行数据清洗，并建立数据透视表。最终生成Excel。

2.运行环境
Python 3.7.8
Selenium 3.13.0

3.依赖包清单

pydivert==2.1.0
PyQt5==5.15.4
qt5-tools==5.15.2.1.2
selenium==3.13.0
requests==2.26.0

pandas == 1.4.3
openpyxl==3.0.9
pywin32==302
XlsxWriter==3.0.2

说明：所需安装依赖包，安装requirements.txt依赖
pip install -r requirements.txt
 文件生成
pip freeze > requirements.txt

4.项目文件的构成
文件名/文件夹名	            说明	                  备注
Investigation.py	     主要爬取数据的功能	    负责爬取企查查的数据
guiRun.py	             设计GUI界面	          用户与界面互动
run.py	               主程序	                运行程序
配置文件/Istry.ini	    配置istry	            保存相关行业的关键词
配置文件/Key.ini	      配置Key	              保存相关关键词
配置文件/Clear_Key.ini	配置Clear_Key	        保存清除相关关键词
配置文件/user.ini	      配置ini	              保存账号和密码
requirements.txt	     项目依赖包	          需要安装模块，可以按依赖包安装
run.spec	             Python打包配置	          run表示是主程序

6.使用说明
1..配置代理ip
如免费代理ip到期，那么所需购买代理ip或者网上搜免费代理ip。
2.参数清单
investigation.py
参数	        类型	    值	              说明
Newfile	      str	  D://....	      保存企查查数据的文件路径
ip/self.ip	  int	  127.0.0.1:80	    代理ip
username	    str	  1899445	      手机号（账号）
password	    str	  ********	      密码
k_dicts	      str	    /	            相关关键词
I_dicts	      str	    /	             相关行业的关键词
c_lists	      str	    /	          清除相关的关键词
rds	          float	  0.80	          相似度去重
filename	    Str	    D://...	    这是读取企查查公司全称名单
data	        str	    -	            保存企查查招标的数据

#PS:注意企查查网站内容变更或者最新的，那么将根据网站去修改相关的模块。
