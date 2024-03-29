# 煲机数据处理工具

---

## 使用说明

- 将log数据放在文件夹下
- 根据提示输入相应信息
- 等待Excel文件生成
- 如果程序不支持该log数据，可以手动修改 log_to_excel.py 程序：
  - 将 log数据中各模块的名称填写进 g_module_name 中
- 若想直接执行源码，第一次需要先运行 [库自动安装脚本.py](./库自动安装脚本.py)，安装相应的Python库文件
- 若log数据开头不是Sample，可修改config.xml文件中begin_value的值 (默认为Sample)
- 可修改config.xml文件中drift_current调整偏移电流大小 (默认100)
  - 只有当电流大于漂移电流值，程序才会判断为充放电阶段
- 可修改config.xml文件中的current_rate调整电流倍数 (默认为1)

---

## 目前通过测试的芯片列表

- BQ40Z50R2
- SN27541-M200
- BQ28Z610
- BQ40Z50R1
- BQ8050

同时也支持列表上没有芯片，只要log数据中模块名相同即可

---

## 版本更新

### V2.1 -- 2023/8/28
1. 修改预估Term过高导致算法不准的问题
2. 添加Max17201的支持

### V2.0 -- 2023/3/9
1. 修改往后时刻点预估Term点的导致软件崩溃的bug
2. 更新空值处理的算法

### V1.5.6(testing) -- 2022/6/29
1. 默认使用EDV为放电截止判断条件

### V1.5.5(testing) -- 2021/11/10
1. 修复首行通讯错误出现的bug

### V1.5.4(testing) -- 2021/10/12
1. 修复fcc为0时程序出现错误的bug
2. 更新库自动安装脚本

### V1.5.3(testing) -- 2021/09/28
1. 添加对 Maxim 17300的支持
2. 添加特殊时间点数据高亮显示
3. 添加恒流和恒功率放电显示

### V1.5.2(testing) -- 2021/09/13
1. 添加电流倍数选项 (在config.xml中修改)
2. 添加Proto阶段输入与显示

### V1.5.1(testing) -- 2021/09/07
1. 添加对bq8050的支持
2. 添加config.xml配置文件，可手动配置开始标记和偏移电流范围

### V1.5.0 -- 2021/09/06
1. 使用xlsxwriter库代替pandas和openpyxl库
2. 生成excel文档速度大幅提高
3. 修复大数据log文件卡死问题

### V1.4.5(testing) -- 2021/09/03
1. 加入对BQ40Z50R1的支持

### V1.4.4(testing) -- 2021/08/30
1. 修复无法从log数据中识别到芯片为bq40z50-R2的bug

### V1.4.3(testing) -- 2021/08/16
1. 添加通过充放电时间检测是否为正确的充放电周期功能
2. 添加未发现term点时，检测往后20个时间点内是否出现term点的功能

### V1.4.2 -- 2021/08/10
1. 修改误差通过范围为6%

### V1.4.1 -- 2021/08/05
1. 更新BQ40Z50R2 term点计算方法

### V1.4.0 -- 2021/07/30
1. 删除自定义模块名功能
2. 更新获取log模块名函数，使其更容易升级维护
3. 加入未发现Term点时的错误处理
4. 加入警告显示功能

### V1.3.4 -- 2021/07/29
1. 增加未找到term_voltage的错误提示

### V1.3.3 -- 2021/07/22
1. 修复电流校正算法中的bug

### V1.3.2 -- 2021/07/20
1. custom_name中增加温度名

### V1.3.1 -- 2021/07/20
1.  重命名代码文件
2.  添加Python库自动安装脚本

### V1.3 -- 2021/07/16
1.  修复错误

### V1.2 -- 2021/07/16
1.  更新通讯错误数据处理算法，修改逻辑错误
2.  增加特殊情况下计算Fuel Gauge Deviation方法
3.  加强对log数据处理的兼容性
4.  bq28z610通过测试

### V1.1 -- 2021/07/16
1.  更新通讯错误数据处理算法

### V1.0 -- 2021/07/16
1.	初始版本

---

:computer: 遇到问题可通过邮箱联系我：

:chestnut: alex.chen@inventuspower.com
