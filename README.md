# Audit Kit
由于通过VBA爬客编表的效率较低，导致公司电脑在爬表过程中死机而影响工作，此 audit_kit Python 版本根据`ABASToolkit.xlam`爬表思路编写。如有 Python 脚本的需求 :+1: 或改进意见 :-1:，请联系作者 [@alfredbowenfeng](https://github.com/alfredbowenfeng)，邮箱是 alfred.bowenfeng@gmail.com。

使用 Python 初始配置可能看起来比较麻烦，实际上只要一次性花5分钟左右的时间安装、配置、安装包，以后爬表的同时，可以用电脑运行其他应用软件，避免死机而无法继续工作的问题。

## Python 的配置
在 Windows 上配置 Python 还另需要设置环境变量、安装包等，因此已经使用 pyinstaller 生成了可直接运行的.exe的应用程序，故电脑不再需要配置 Python 。

## Breakdown by Value
这里的“Value”指的是每个单元格的内容，可以是数值，也可以是文字、日期等。

### EXE
如果你直接使用.exe的应用程序，那么在开始会要求你输入配置，请根据提示与例子输入并运行即可。

### Jupyter Notebook
如果你使用 [Anaconda](https://www.anaconda.com/download/) 的 Jupyter Notebook 来调试或使用，请直接修改`breakdown_value.ipynb`中第一格的配置即可。

## Breakdown by Link
这里的“Link”指的是每个单元格的链接，在 Excel 内可选中单元格，点击`ctrl + {/[`便可查看链接指向。由于链接是程序自动生成的，故在使用 Excel 打开生成文件后，需要全选、复制、粘贴才可生效。

### EXE
如果你直接使用.exe的应用程序，那么在开始会要求你输入配置，请根据提示与例子输入并运行即可。

### Jupyter Notebook
如果你使用 [Anaconda](https://www.anaconda.com/download/) 的 Jupyter Notebook 来调试或使用，请直接修改`breakdown_link.ipynb`中第一格的配置即可。