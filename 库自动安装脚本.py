import os

print("自动安装Python依赖库，请等待......")
if os.system("pip install openpyxl"):
    pass
    
if os.system("pip install pandas"):
    pass
    
input("安装完成，按任意键退出")