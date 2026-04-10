# 脚轮测试上位机系统

![脚轮logo](./img.png)

## 项目概述

这是一个基于TCP通信的脚轮测试上位机系统，用于实时监测脚轮测试过程中的各种参数，并生成测试报表。系统通过Modbus TCP协议与PLC进行通信，读取测试数据并在界面上实时显示。

## 功能特点

### 🎛️ 参数界面
- **产品信息管理**：输入产品名称、型号、编号、轮径、材质、硬度等信息
- **测试参数设置**：配置测试里程、速度、时长、间隔时间、障碍数量、障碍高度、承载重量、承载温度等参数
- **实时数据监测**：显示速度、循环时间、测试时间、里程、温度、压力等实时数据
- **手动控制功能**：转盘正转/反转、砝码测量、数据清除等点动控制
- **PLC通信**：将测试参数导入到PLC，启动/停止测试

### 📈 数据曲线
- **实时温度监测**：工位1和工位2的温度变化曲线
- **实时压力监测**：工位1和工位2的压力变化曲线
- **数据可视化**：30秒/点的采样频率，最多保存2880个数据点

### 📋 测试报表
- **报表预览**：实时查看测试数据快照
- **Excel导出**：生成详细的脚轮动力检验报告
- **数据记录**：包含产品信息、测试条件、实测数据等完整信息

## 技术栈

- **开发语言**：Python 3.x
- **GUI框架**：PyQt5
- **数据可视化**：pyqtgraph
- **Modbus通信**：pymodbus
- **Excel处理**：openpyxl
- **数据处理**：numpy

## 安装步骤

1. **克隆项目**
   ```bash
   git clone <repository-url>
   cd caster
   ```

2. **安装依赖**
   ```bash
   pip install -r requirements.txt
   ```

   或者手动安装所需库：
   ```bash
   pip install PyQt5 pyqtgraph pymodbus openpyxl numpy
   ```

3. **准备资源文件**
   - 确保 `jiaolun.png` 图片文件存在于项目根目录

## 使用说明

### 基本操作流程

1. **启动应用**
   ```bash
   python main.py
   ```

2. **配置产品信息**
   - 在"产品信息"区域填写相关信息
   - 系统会自动保存这些信息，下次启动时会自动加载

3. **设置测试参数**
   - 在"参数设置"区域输入测试参数
   - 确保所有参数都已正确填写

4. **导入参数到PLC**
   - 点击"⬇️ 导入方案到 PLC"按钮
   - 系统会将参数写入PLC寄存器

5. **启动测试**
   - 点击"▶️ 启动"按钮开始测试
   - 实时数据会在界面上显示

6. **查看数据曲线**
   - 切换到"📈 数据曲线"标签页
   - 查看温度和压力的实时变化曲线

7. **生成测试报表**
   - 点击"🖨️ 生成报表"按钮
   - 选择保存位置，系统会生成Excel格式的测试报告

### 手动控制功能

- **转盘正转/反转**：点动控制转盘旋转方向
- **砝码1/2测量**：开启/关闭砝码测量功能
- **数据清除**：清除测试数据
- **测试报表预览**：查看当前测试数据的快照

## PLC通信

- **通信协议**：Modbus TCP
- **默认IP地址**：192.168.6.6
- **默认端口**：502
- **数据寄存器**：读取D0-D46，写入D12、D14、D16、D46、D48、D50等

## 系统要求

- Windows 7/10/11
- Python 3.6+
- 4GB以上内存
- 1024x768以上分辨率

## 注意事项

1. 确保PLC与上位机在同一网络中
2. 导入参数前确保所有参数已正确填写
3. 生成报表前建议先预览数据快照
4. 如遇到连接问题，请检查网络连接和PLC状态

## 技术支持

技术支持：康瑞智能化科技有限公司

## 版本信息

- 版本：1.0.0
- 最后更新：2026-03-31

## 打包教程

```bash
nuitka --standalone --windows-disable-console --enable-plugin=pyqt5 --include-data-files=jiaolun.png=jiaolun.png main.py
```
exe同目录下必须包含libopenblas64__v0.3.23-293-gc2f4bdbb-gcc_10_3_0-2bde3a66a51006b2b53eb373ff767a3f.dll

带图标打包：
```bash
nuitka --standalone --windows-disable-console --windows-icon-from-ico=star.ico --enable-plugin=pyqt5 --include-package-data=qfluentwidgets --include-data-files=jiaolun.png=jiaolun.png --include-data-files=star.ico=star.ico main.py
```
