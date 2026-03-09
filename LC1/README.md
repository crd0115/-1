# LC1 - BOM Assistant（智能物料清单生成器）

一个智能的物料清单（BOM）生成工具，可以识别电路图中的电子元器件，自动与立创商城联动查询物料，生成完整的BOM表。

## 功能特性

- **电路图识别**：上传电路图，自动识别电路中的电子元器件
- **元器件分析**：提取元器件的参考标号、参数、封装等信息
- **立创商城联动**：自动查询立创商城中对应的物料和价格
- **BOM表生成**：生成标准的物料清单表，支持CSV导出
- **库存查询**：实时查询物料的库存信息

## Installation

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Required libraries:
   - pillow - 图像处理
   - opencv-python - 图像识别
   - numpy - 数值计算
   - requests - HTTP请求
   - pandas - 数据处理
   - easyocr - 光学字符识别(OCR)
   - torch/torchvision - 深度学习框架

## Usage

Run the main program:
```bash
python main.py
```

Then follow these steps:
1. Enter the path to your circuit diagram image
2. The program will analyze the circuit and extract components
3. It will query the LiChuang marketplace for each component
4. A BOM table will be generated and saved to file

## File Structure

```
LC1/
├── main.py - Main program
├── requirements.txt - Dependencies
├── README.md - This file
├── .github/
│   └── copilot-instructions.md - Project instructions
└── output/
    └── BOM.csv - Generated BOM file
```

## Supported Image Formats

- JPEG (.jpg, .jpeg)
- PNG (.png)
- BMP (.bmp)
- TIFF (.tiff)

## BOM Output Format

The generated BOM table includes:
- Reference - Element reference (R1, C1, U1, etc.)
- Value - Component value
- Package - Physical package type
- Part Number - Catalog number from LiChuang
- Quantity - Number of units
- Supplier - Supplier name
- Unit Price - Price per unit
- Total Price - Total cost for all units

## Future Enhancements

- Integration with more suppliers
- PDF circuit diagram support
- Automatic component matching algorithm
- Price comparison across suppliers
- Real-time inventory tracking
- Component alternative suggestions

## Requirements

- Python 3.7+
- Google Chrome (for potential browser automation)

## Note

This project requires valid API access to LiChuang marketplace for full functionality.