# WOS_crawler

## 简介

利用 selenium 自动化控制 Chrome 浏览器以 Excel 格式导出 Web of Science 搜索结果。

目前暂时只支持导出文献标题、作者、摘要这三个信息，需要的话可在 `WOS.py` 中 `exportExcel()` 中更改。未来会加入选择导出信息功能（咕咕咕？



## 配置环境

此脚本使用需要安装 Chrome 浏览器和 ChromeDriver。若你已经下载过 ChromeDriver 并将其添加到系统路径后，可跳过这一步。如果你下载的 ChromeDriver 并未添加到系统路径，可自行修改 `driver.py` 中的 `ChromeDriver_PATH`

先下载 Chrome 浏览器，安装完成后打开浏览器，在 设置->关于 中找到浏览器的版本信息，记下 Chrome  的版本号，然后打开 ChromeDriver 的下载网站。[ChromeDriver 下载链接](https://chromedriver.chromium.org/downloads)

下载 ChromeDriver 时需要注意下载的版本，ChromeDriver 的版本要与 Chrome 浏览器的版本一样。如果实在找不到一样的，就下载与浏览器版本最接近的。

下载文件时注意选择与操作系统相匹配的版本，下载完成后解压，将解压出的 `chromedriver` 文件放到脚本相同的目录下。



## 运行

脚本运行前可使用文本编辑器打开 `main.py`，在开头中填入**搜索结果网址**、**文件下载存放路径**、**登录用学号姓名**等信息，填写后可提升脚本自动化程度。以上信息可以不填，不影响脚本导出功能。

运行脚本需要 python3.7 及以上版本，低版本的适配性未知，不保证脚本一定能顺利运行。

使用前先安装依赖，在脚本目录下打开命令行后输入 `python -m pip install -r requirements.txt ` 安装所需依赖。

安装完成后直接运行 `main.py` 即可。



## 注意事项

脚本运行需要一定的网络条件，网络条件差会影响脚本运行。网络条件较差的情况下运行脚本可适当增大 `main.py` 中 `RETRY` 的数值。
