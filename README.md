# Reactivity-test-software
这是一个由PyQt5实现的反应力测试程序，旨在测试人眼在不同色彩对比度的情况下，信息搜索能力的差异

humanfactor.py文件可以直接执行，也可以用
`pyinstaller --onefile --noconsole --add-data "image.png;." humanfactor.py`
来编译成exe程序

运行后，会自动生成`results.xlsx`文件，会记录用户的准确率和测试时间
