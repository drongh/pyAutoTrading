# PyAutoTrading
股票交易软件辅助工具

## 简介
用于华泰证券独立交易软件，同时支持通达信版本和同花顺版本，以上版本都需有`双向委托`界面。软件可以一次监控5只股票，根据条件下单。每次下单耗时小于1s。
如果有疑问，或是建议，可以发邮件联系，邮箱：`ronghui.ding@outlook.com`，或加入QQ群：`486224275`。
开发环境:win10 64bit, python3 64bit、pywin32、tushare。 以前是用python 32bit开发的，现在好像python 64bit的也能用。

## 同花顺版使用说明
* 华泰的最好用通用版，稳定性高，行情可以选用同花顺的行情，行情速度快了5s，对于追涨杀跌挺重要的。
* 软件共有3个文件，`pyautotrade_ths.pyw`主程序，`stockInfo.dat`存盘文件，`winguiauto.py`是封装的winapi函数。
* 交易软件启动后，按`F6`，进入`双向委托`界面，启动本程序后，不要再切换交易软件的操作界面。切换到其他界面后即使再切换回来，会不能正常获取句柄，无法下单。
* 不写时间条件单，默认时间为凌晨1点。时间条件满足后才检查价格条件，如果只想要时间条件单而忽略价格条件单，可以写个始终满足条件的价格。
* 股票数量最好为100的倍数，小于100股的不会交易，大于100、非整数倍的将取整, 比如150股将作为100股。
* 时间为24小时制，形式为 `时：分：秒`， 每项都必须写， 后面的写法是错误的： `13：30`。
* 同花顺版本支持获取持仓情况，需要把持仓表单设定为11列，通过设定`是否可见`,请看图。获取持仓数据时需要通过鼠标获取焦点(这是自动的),最小化的时会自动从任务栏恢复。
另外，在交易软件上，把`持仓`表单显示出来，不要放在`成交`或`其他表单`。
* 想使用同花顺通用版的，在`pyautotrade_ths.pyw`文件中`Operation`类的`__init__`函数的第4或5行选择一行就可以。

## 通达信版使用说明
* 软件共有3个文件，`pyautotrade_tdx.pyw`主程序，`stockInfo.dat`存盘文件，`winguiauto.py`是封装的winapi函数
* 交易软件启动后，直接点击左边树形列表`对买对卖`，然后启动本程序，不要再切换到其它界面，始终在`对买对卖`界面上
* 不写时间条件单，默认时间为凌晨1点。时间条件满足后才检查价格条件，如果只想要时间条件单而忽略价格条件单，可以写个始终满足条件的价格。
* 股票数量为100的倍数, 如果输入150股将作为100股。默认为0股，也就说，股票数量由交易软件自动填写，当然需提前在交易软件里设定好，在`系统设置-仓位策略`里选择固定数量, 见图5
* 时间为24小时制，形式为 `时：分：秒`， 每项都必须写， 后面的写法是错误的： `13：30`
* 交易软件的委托价格由交易软件自动填写，在`系统设置-自动策略`， 启用`启用自动跟盘`， 自己决定选哪个价格
* 自动交易时最好最大化窗口，否则不能获得持仓情况

## 版本
* v 0.01 修正了股票价格实时显示问题。
* v 0.02 重构了交易软件接口，目前在最小化状态下也可以下单，下单速度加快，增加委托日志。
* v 0.03 重新布局了控件，修改委托日志控件。修复了少许Bug。
* v 0.04 重新布局了控件，重构了monitor函数。现在一次可以下4个条件单。
* v 0.05 加入时间条件单。
* v 0.06 交易软件接口函数单独放winguiauto文件。
* v 0.07 时间条件单和价格条件单相结合，添加保存和载入功能,存档和主文件在同一目录下，名为stockInfo.dat，是个二进制文件。
* v 0.08 代码清理，添加了注释。现在可以同时监控5只股票。
* v 0.09 加入自动刷新功能，每隔5分钟刷新一次，防止软件进入待机状态。
* v 0.10 修改了几个bug，买卖价格改由python计算，加快了下单速度（0.6s），稳定性增加了不少。需更改交易软件设置，请看图。
* v 0.11 同花顺版修复了一个严重的bug。
* v 0.12 接口函数改用类重写。winguiauto.py和通达信一致, 修复一个UI显示错误。
* v 0.13 同花顺版本支持获取持仓和资金了。通达信版没有变化。
* v 0.14 增加通达信版改进版。此版用pywinauto改写。大幅简化代码。
## Bugs
* 集合竞价时获取价格为零,导致误下单



-----------------------------------

![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting1_ths.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting2_ths.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting3_ths.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting4_ths.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting5_ths.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting6_ths.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting1_tdx.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting2_tdx.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting3_tdx.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting4_tdx.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting5_tdx.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting6_tdx.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/setting7_tdx.png)
![image](https://github.com/drongh/PyAutoTrading/raw/master/Logo/trading_ths.png)
