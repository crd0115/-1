# BOM Assistant - 使用指南

**智能物料清单生成器 + 立创商城集成**

## 功能概述

这个工具可以：
1. ✅ 从电路图图像中提取元器件信息（使用 EasyOCR）
2. ✅ 支持手动输入元器件清单
3. ✅ 连接立创商城数据库查询价格和库存
4. ✅ 自动合并相同的元器件（统计数量）
5. ✅ 生成详细的 BOM 表（含价格、库存、总成本）
6. ✅ 导出为 CSV 格式（可在 Excel 中打开）

---

## 使用方式

### 方式 1: 从电路图图像分析（自动识别）

```bash
python main.py
```

选择选项 `1`，输入电路图图像路径。程序会：
- 使用 EasyOCR 自动识别元器件标签（例如 R1、C1）
- 将识别出的标签打印出来
- **接着提示你为每个标签手动输入元器件值和封装**（例如 `10k,0603`）
- 使用这些输入构建组件清单并自动查询立创商城价格

如果OCR检测失败或图像不清晰，程序会退回到示例列表。

**优点**：根据实际图片提取标签，避免写死数据
**缺点**：需要用户按提示补充每个元件的规格，OCR偶尔会漏检标签。

---

### 方式 2: 使用示例数据（快速测试）

```bash
python main.py
```

选择选项 `2`，程序会加载预定义的示例元器件：
- 16 x 0.1µF 电容（0603 封装）
- 1 x LM1117 稳压芯片（SOT-223）
- 1 x 40MHz 晶振

**适用场景**：演示、测试或快速验证流程

---

### 方式 3: 手动输入清单（精确控制）

```bash
python main.py
```

选择选项 `3`，按照提示输入每个元器件：

```
Component (e.g., R1|10k|0603): R1|10k|0603
✓ Added R1

Component (e.g., R1|10k|0603): C1|100n|0603
✓ Added C1

Component (e.g., R1|10k|0603): done
```

**优点**：精确性最高，完全控制
**缺点**：需要手动输入（对于大电路可能较费时）

---

## 输出文件

### BOM 表 CSV 格式

生成的 `BOM.csv` 或 `BOM_output.csv` 包含以下列：

| 列 | 说明 |
|---|---|
| Reference | 元器件参考编号（如 R1, C1, U1） |
| Value | 元器件规格值（如 10k, 100n, STM32F103） |
| Package | 封装类型（如 0603, SOT-223, LQFP48） |
| Quantity | 数量（自动合并相同元器件后的统计）|
| Part Number | 立创商城零件编号 |
| Supplier | 供应商（立创商城） |
| Unit Price (¥) | 单价（元） |
| Total Price (¥) | 总价（元） = 单价 × 数量 |

---

## 示例输出

运行示例后的输出：

```
============================== BOM 表 - 物料清单 ==============================

Reference                                                 Value    Package  Quantity  Part Number          Supplier  Unit Price (¥)  Total Price (¥)
C5,C6,C7,C8,C9,C10,C11,C12,C13,C14,C15,C16,C17,C30,C31,C32  0.1µF     0603       16    C0603X5R1V104K030BC   立创商城      0.0200            0.32
U2                                                        LM1117   SOT-223        1    LM1117-3.3            立创商城      0.5000            0.50
X1                                                        40MHz     DIP          1    HC-49US-40M           立创商城      0.3000            0.30

================================================================================
总成本: ¥1.12
```

---

## 代码架构

### 核心类

#### 1. **CircuitImageAnalyzer** - 电路图分析
```python
analyzer = CircuitImageAnalyzer()
components = analyzer.extract_component_info("circuit.png")
```
- `load_image()`: 加载电路图图像
- `detect_component_labels()`: 使用 EasyOCR 识别文本标号
- `extract_component_info()`: 提取元器件信息

#### 2. **LiChuangAPI** - 立创商城集成
```python
api = LiChuangAPI(api_key="your_api_key")
results = api.search_component("0.1µF", "0603")
```
- `search_component()`: 搜索元器件
- `get_component_details()`: 获取详细信息
- 本地数据库（可扩展为真实 API 调用）

#### 3. **BOMGenerator** - BOM 表生成
```python
bom_gen = BOMGenerator()
bom_df = bom_gen.generate_bom(components, merge=True)
bom_gen.save_bom(bom_df, "BOM.csv")
```
- `add_component()`: 添加元器件
- `merge_components()`: 合并相同元器件
- `search_and_enrich()`: 查询价格和库存
- `generate_bom()`: 生成 Pandas DataFrame
- `save_bom()`: 保存为 CSV
- `display_bom()`: 打印表格

---

## 扩展功能

### 集成真实的立创商城 API

将 `LiChuangAPI` 中的 `_init_component_database()` 替换为真实 HTTP 请求：

```python
def search_component(self, component_name: str, package: str = "") -> List[Dict]:
    """搜索元器件"""
    url = f"{self.base_url}/search"
    params = {
        "keyword": component_name,
        "package": package,
        "apiKey": self.api_key
    }
    response = requests.get(url, params=params)
    return response.json()["results"]
```

---

## 常见问题

**Q: 如何整合我自己的电路图？**
- 将电路图保存为 PNG/JPG 格式
- 运行程序，选择选项 1
- 输入图像路径

**Q: 如何获得立创商城 API 密钥？**
- 访问 https://www.iczoom.com
- 注册账户并申请 API 权限
- 在代码中配置: `api = LiChuangAPI(api_key="your_key")`

**Q: 合并元器件是什么意思？**
- 程序会统计相同规格的元器件（按 Value+Package）
- 例如：C5, C6, C7 都是 0.1µF 0603 → 合并为 16x 0.1µF 0603
- 这样便于批量采购

**Q: 为什么价格显示 0.0？**
- 当前数据库中没有某个元器件的信息
- 需要更新立创商城数据库或连接真实 API

---

## 依赖项

```
pillow           # 图像处理
opencv-python    # 图像加载
numpy            # 数值运算
requests         # HTTP 请求
pandas           # 数据处理和 CSV 导出
easyocr          # 光学字符识别（OCR）
torch            # 深度学习（EasyOCR 依赖）
torchvision      # 计算机视觉（EasyOCR 依赖）
```

安装依赖：
```bash
pip install -r requirements.txt
```

---

## 下一步

- [ ] 集成真实的立创商城 API
- [ ] 支持导出到 Excel（含格式化）
- [ ] 添加元器件的数据手册链接
- [ ] 支持多供应商价格比较
- [ ] 实现库存预警功能
- [ ] 添加成本分析仪表板

---

## 联系支持

如有问题或建议，请提交 Issue 或联系项目维护者。

**最后更新**: 2026年3月8日
