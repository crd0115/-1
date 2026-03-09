"""
BOM Assistant - 智能物料清单生成器
流程：OCR 识别 → AI 补充 → Excel 输出
"""

import os
import json
import base64
import logging
import webbrowser
import re
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd
import requests
import dashscope
from bs4 import BeautifulSoup

# OCR 库
try:
    import easyocr
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    easyocr = None

try:
    import pytesseract
    from PIL import Image
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False
    pytesseract = None
    Image = None

# Excel 支持
try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    openpyxl = None

# ==================== 配置管理 ====================

@dataclass
class Config:
    """配置管理类"""
    qwen_api_key: str = ""
    lichuang_api_key: str = ""
    project_name: str = "BOM_Project"
    ocr_language: str = "ch_sim,en"
    log_level: int = logging.INFO
    output_format: str = "excel"  # excel 或 csv
    
    @classmethod
    def from_env(cls) -> "Config":
        """从环境变量加载配置"""
        return cls(
            qwen_api_key=os.getenv("QWEN_API_KEY", ""),
            lichuang_api_key=os.getenv("LICHUANG_API_KEY", ""),
        )
    
    def save_to_file(self, path: str = "config.json"):
        """保存配置到文件"""
        with open(path, "w", encoding="utf-8") as f:
            json.dump({
                "lichuang_api_key": self.lichuang_api_key,
                "output_format": self.output_format
            }, f, indent=2)
    
    @classmethod
    def load_from_file(cls, path: str = "config.json") -> "Config":
        """从文件加载配置"""
        config = cls.from_env()
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                config.lichuang_api_key = data.get("lichuang_api_key", "")
                config.output_format = data.get("output_format", "excel")
        return config


# ==================== 日志设置 ====================

def setup_logging(log_level: int = logging.INFO, log_file: str = "") -> logging.Logger:
    """设置日志系统"""
    logger = logging.getLogger("BOM_Assistant")
    logger.setLevel(log_level)
    
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        
        if log_file:
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
    
    return logger


logger = setup_logging()


# ==================== 自定义异常 ====================

class BOMError(Exception):
    """BOM 生成器基础异常"""
    pass


class OCR_Error(BOMError):
    """OCR 识别异常"""
    pass


class AI_Error(BOMError):
    """AI 识别异常"""
    pass


# ==================== 工具函数 ====================

def create_project_folder(folder_name: str = "BOM_Project") -> Optional[Path]:
    """在桌面创建项目文件夹"""
    try:
        desktop_path = Path.home() / "Desktop"
        project_path = desktop_path / folder_name
        project_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"✓ 项目文件夹：{project_path}")
        return project_path
    except Exception as e:
        logger.error(f"✗ 创建文件夹失败：{e}")
        return None


def select_image_file() -> Optional[str]:
    """选择图像文件"""
    path = input("请输入图像文件路径 (留空弹出文件选择器): ").strip()
    if path:
        return path if os.path.exists(path) else None
    
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        path = filedialog.askopenfilename(
            title="选择电路图像文件",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.tiff"), ("All Files", "*")]
        )
        return path if path and os.path.exists(path) else None
    except Exception as e:
        logger.error(f"文件选择器失败：{e}")
        return None


def select_save_file(default_name: str = "BOM.xlsx") -> Optional[str]:
    """选择保存文件路径"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        file_type = [("Excel 文件", "*.xlsx")] if EXCEL_AVAILABLE else [("CSV 文件", "*.csv")]
        path = filedialog.asksaveasfilename(
            title="保存 BOM 表",
            defaultextension=".xlsx" if EXCEL_AVAILABLE else ".csv",
            initialdir=Path.home() / "Desktop",
            filetypes=file_type
        )
        return path if path else None
    except Exception as e:
        logger.error(f"保存对话框失败：{e}")
        return None


# ==================== OCR 识别模块 ====================

class OCRRecognizer:
    """OCR 图像识别器"""
    
    def __init__(self, languages: str = "ch_sim,en"):
        self.languages = languages.split(",")
        self.reader = None
        self._init_ocr()
    
    def _init_ocr(self):
        """初始化 OCR 引擎"""
        if OCR_AVAILABLE and easyocr:
            try:
                self.reader = easyocr.Reader(self.languages, gpu=False)
                logger.info("✓ EasyOCR 初始化成功")
            except Exception as e:
                logger.warning(f"EasyOCR 初始化失败：{e}")
        elif TESSERACT_AVAILABLE and pytesseract and Image:
            logger.info("✓ 使用 Tesseract OCR")
        else:
            logger.warning("⚠ 未找到可用的 OCR 引擎，将直接使用 AI 识别")
    
    def recognize_text(self, image_path: str) -> str:
        """从图像中提取文本"""
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"图像文件不存在：{image_path}")
        
        text = ""
        
        # 优先使用 EasyOCR
        if self.reader:
            try:
                results = self.reader.readtext(image_path)
                text = "\n".join([result[1] for result in results])
                logger.info(f"EasyOCR 提取 {len(results)} 个文本块")
            except Exception as e:
                logger.warning(f"EasyOCR 识别失败：{e}")
        
        # 回退到 Tesseract
        if not text and TESSERACT_AVAILABLE and pytesseract and Image:
            try:
                img = Image.open(image_path)
                text = pytesseract.image_to_string(img, lang='chi_sim+eng')
                logger.info("Tesseract 识别完成")
            except Exception as e:
                logger.warning(f"Tesseract 识别失败：{e}")
        
        return text
    
    def extract_components_from_text(self, text: str) -> List[Dict[str, str]]:
        """从 OCR 文本中提取元器件信息"""
        components = []
        
        # 元器件识别正则模式
        patterns = [
            # 电阻：R1 10k 0603
            r'(R\d+)[\s:]+([\d\.]+[kKmMΩΩ]?)[\s]+(\d+)',
            # 电容：C1 10uF 0805
            r'(C\d+)[\s:]+([\d\.]+[uUnNpPfF]?)\s+(\d+)',
            # 通用：参考编号 值 封装
            r'([RCULDQ]\d+)[\s:]+([\w\.]+)[\s]+(\d+)',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                if len(match) >= 3:
                    components.append({
                        "reference": match[0].strip(),
                        "value": match[1].strip(),
                        "package": match[2].strip(),
                        "source": "OCR"
                    })
        
        # 去重
        seen = set()
        unique_components = []
        for comp in components:
            key = f"{comp['reference']}_{comp['value']}"
            if key not in seen:
                seen.add(key)
                unique_components.append(comp)
        
        logger.info(f"OCR 提取 {len(unique_components)} 个元器件")
        return unique_components


# ==================== AI 识别模块 ====================

class AIRecognizer:
    """AI 图像识别器（千问大模型）"""
    
    def __init__(self, api_key: str):
        if not api_key.strip():
            raise ValueError("Qwen API Key 不能为空")
        self.api_key = api_key
        dashscope.api_key = api_key
        logger.info("✓ AI 识别器初始化完成")
    
    def load_image_base64(self, image_path: str) -> str:
        """加载图像并编码为 base64"""
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    
    def recognize_components(self, image_path: str, ocr_results: List[Dict] = None) -> List[Dict[str, str]]:
        """使用 AI 识别电路图中的元器件"""
        try:
            logger.info("正在编码图像...")
            image_base64 = self.load_image_base64(image_path)
            
            # 构建提示词，包含 OCR 结果作为参考
            prompt = self._build_prompt(ocr_results)
            
            logger.info("正在调用 Qwen AI 进行元件识别...")
            response = dashscope.MultiModalConversation.call(
                model='qwen-image-2.0-pro',
                messages=[{
                    'role': 'user',
                    'content': [
                        {'image': f'data:image/png;base64,{image_base64}'},
                        {'text': prompt}
                    ]
                }]
            )
            
            if response.status_code == 200:
                content = response.output.choices[0].message.content
                logger.info(f"AI 响应长度：{len(content)}")
                components = self._parse_ai_response(content)
                if components:
                    logger.info(f"✓ AI 识别 {len(components)} 个元器件")
                    return components
            else:
                logger.error(f"AI API 失败：{response.status_code}")
        
        except Exception as e:
            logger.error(f"AI 识别失败：{e}")
        
        return []
    
    def _build_prompt(self, ocr_results: List[Dict] = None) -> str:
        """构建 AI 提示词"""
        base_prompt = """分析这张电路图，识别所有电子元件。
输出为 JSON 格式的列表，每个元件包含：
- reference：参考编号（如 R1、C1、U1）
- value：值（如 10k、10uF、STM32F103C8T6）
- package：封装（如 0603、0805、LQFP48）
- type：元件类型（如 Resistor、Capacitor、IC）

只输出 JSON 数组，不要其他文字。"""

        if ocr_results:
            ocr_text = "\n".join([f"{c['reference']} {c['value']} {c['package']}" for c in ocr_results[:10]])
            base_prompt += f"\n\n以下是 OCR 已识别的部分元件（仅供参考和补充）：\n{ocr_text}"
            base_prompt += "\n请补充 OCR 未识别或识别不完整的元件。"
        
        return base_prompt
    
    def _parse_ai_response(self, content: str) -> List[Dict[str, str]]:
        """解析 AI 响应"""
        try:
            # 清理可能的 markdown 标记
            content = re.sub(r'```json\s*|\s*```', '', content.strip())
            components = json.loads(content)
            
            if isinstance(components, list):
                for comp in components:
                    comp['source'] = 'AI'
                    # 标准化字段
                    if 'reference' not in comp:
                        comp['reference'] = comp.get('ref', comp.get('name', ''))
                    if 'value' not in comp:
                        comp['value'] = comp.get('val', '')
                    if 'package' not in comp:
                        comp['package'] = comp.get('pkg', comp.get('footprint', ''))
                return components
        except json.JSONDecodeError as e:
            logger.warning(f"JSON 解析失败：{e}")
        
        # 尝试从文本中提取
        return self._extract_from_text(content)
    
    def _extract_from_text(self, text: str) -> List[Dict[str, str]]:
        """从文本中提取元器件信息"""
        components = []
        patterns = [
            r'["\']?reference["\']?\s*[:\s]\s*["\']?([RCULDQ]\d+)["\']?.*?["\']?value["\']?\s*[:\s]\s*["\']?([\w\.]+)["\']?.*?["\']?package["\']?\s*[:\s]\s*["\']?(\d+)["\']?',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE | re.DOTALL)
            for match in matches:
                if len(match) >= 3:
                    components.append({
                        "reference": match[0].strip(),
                        "value": match[1].strip(),
                        "package": match[2].strip(),
                        "source": "AI"
                    })
        
        return components


# ==================== 立创商城 API ====================

class LiChuangAPI:
    """立创商城 API 集成类"""
    
    def __init__(self, api_key: str = ""):
        self.api_key = api_key
        self.base_url = "https://api.lceda.cn"
        self.use_real_api = bool(api_key.strip())
        self.component_database = self._init_component_database()
        logger.info(f"LiChuangAPI 初始化，使用真实 API: {self.use_real_api}")
    
    def _init_component_database(self) -> Dict[str, Dict[str, Any]]:
        """初始化本地元器件数据库"""
        return {
            "0.1uF": {"part_number": "C0603X5R1V104K030BC", "name": "贴片电容 0.1uF 50V X7R 0603", "price": 0.02, "stock": 50000},
            "10uF": {"part_number": "C0805X5R1V106K030BC", "name": "贴片电容 10uF 25V X5R 0805", "price": 0.05, "stock": 20000},
            "10k": {"part_number": "R0603-10K", "name": "贴片电阻 10kΩ 1% 0603", "price": 0.005, "stock": 100000},
            "1k": {"part_number": "R0603-1K", "name": "贴片电阻 1kΩ 1% 0603", "price": 0.005, "stock": 100000},
            "STM32F103C8T6": {"part_number": "STM32F103C8T6-LQFP48", "name": "STM32F103C8T6 微控制器 LQFP48", "price": 8.5, "stock": 500},
            "LM1117-3.3": {"part_number": "LM1117-3.3-SOT223", "name": "LM1117-3.3 线性稳压芯片 SOT-223", "price": 0.5, "stock": 1000},
            "USB-B": {"part_number": "USB-B-SMD", "name": "USB-B 接口 SMD", "price": 0.8, "stock": 2000},
        }
    
    def search_component(self, component_name: str, package: str = "") -> List[Dict[str, Any]]:
        """搜索元器件"""
        logger.debug(f"搜索：{component_name}")
        
        # 本地数据库搜索
        search_results = []
        for key, comp_data in self.component_database.items():
            if component_name.lower() in key.lower() or component_name.lower() in comp_data["name"].lower():
                search_results.append({
                    "part_number": comp_data["part_number"],
                    "name": comp_data["name"],
                    "price": comp_data["price"],
                    "stock": comp_data["stock"],
                    "supplier": "立创商城",
                    "datasheet": comp_data.get("datasheet", "")
                })
        
        if not search_results:
            search_results = [{
                "part_number": f"UNKNOWN-{component_name}",
                "name": f"{component_name} (需手动确认)",
                "price": 0.0,
                "stock": 0,
                "supplier": "立创商城",
                "datasheet": ""
            }]
        
        return search_results


# ==================== BOM 生成器 ====================

class BOMGenerator:
    """BOM 表生成类"""
    
    def __init__(self, api_key: str = "", output_format: str = "excel"):
        self.components: List[Dict[str, Any]] = []
        self.lichuang_api = LiChuangAPI(api_key)
        self.output_format = output_format
        logger.info(f"BOMGenerator 初始化，输出格式：{output_format}")
    
    def add_components(self, components: List[Dict[str, Any]]):
        """添加元器件列表"""
        for comp in components:
            comp.setdefault("quantity", 1)
            self.components.append(comp)
        logger.info(f"已添加 {len(components)} 个元器件")
    
    def merge_components(self, components: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """合并相同的元器件"""
        merged: Dict[str, Dict[str, Any]] = {}
        for comp in components:
            key = f"{comp['value']}_{comp['package']}"
            if key not in merged:
                merged[key] = {
                    "value": comp["value"],
                    "package": comp["package"],
                    "references": [comp["reference"]],
                    "quantity": 1,
                    "part_number": comp.get("part_number", ""),
                    "type": comp.get("type", ""),
                    "source": comp.get("source", "")
                }
            else:
                merged[key]["references"].append(comp["reference"])
                merged[key]["quantity"] += 1
        return list(merged.values())
    
    def search_and_enrich(self, components: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """搜索并丰富元器件信息"""
        for comp in components:
            search_results = self.lichuang_api.search_component(comp["value"], comp.get("package", ""))
            if search_results:
                best_match = search_results[0]
                comp["part_number"] = best_match.get("part_number", "")
                comp["price"] = best_match.get("price", 0)
                comp["stock"] = best_match.get("stock", 0)
                comp["supplier"] = best_match.get("supplier", "立创商城")
        return components
    
    def generate_bom(self, merge: bool = True) -> pd.DataFrame:
        """生成 BOM 表"""
        components = self.components.copy()
        
        if merge:
            components = self.merge_components(components)
        
        components = self.search_and_enrich(components)
        
        bom_data = []
        for comp in components:
            references = ",".join(comp.get("references", [comp.get("reference", "")]))
            unit_price = float(comp.get("price", 0))
            quantity = int(comp.get("quantity", 1))
            total_price = unit_price * quantity
            
            bom_data.append({
                "序号": len(bom_data) + 1,
                "参考编号": references,
                "元件类型": comp.get("type", self._guess_type(comp.get("reference", ""))),
                "值/型号": comp.get("value", ""),
                "封装": comp.get("package", ""),
                "数量": quantity,
                "立创编号": comp.get("part_number", ""),
                "供应商": comp.get("supplier", "立创商城"),
                "单价 (CNY)": f"{unit_price:.4f}",
                "总价 (CNY)": f"{total_price:.2f}",
                "库存": comp.get("stock", 0),
                "识别来源": comp.get("source", "手动")
            })
        
        return pd.DataFrame(bom_data)
    
    def _guess_type(self, reference: str) -> str:
        """根据参考编号猜测元件类型"""
        if not reference:
            return "Unknown"
        prefix = reference[0].upper()
        type_map = {
            'R': '电阻',
            'C': '电容',
            'L': '电感',
            'D': '二极管',
            'Q': '三极管/MOS',
            'U': 'IC 芯片',
            'J': '连接器',
            'F': '保险丝',
            'S': '开关'
        }
        return type_map.get(prefix, "其他")
    
    def save_bom(self, bom_df: pd.DataFrame, output_path: str) -> bool:
        """保存 BOM 表（支持 Excel 和 CSV）"""
        try:
            if self.output_format == "excel" and EXCEL_AVAILABLE:
                # 确保扩展名为 .xlsx
                if not output_path.endswith('.xlsx'):
                    output_path = output_path.rsplit('.', 1)[0] + '.xlsx'
                
                # 使用 Excel 写入器添加格式
                with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                    bom_df.to_excel(writer, index=False, sheet_name='BOM 表')
                    
                    # 获取 workbook 和 worksheet
                    workbook = writer.book
                    worksheet = writer.sheets['BOM 表']
                    
                    # 设置列宽
                    column_widths = [10, 15, 12, 20, 12, 10, 18, 15, 12, 12, 10, 12]
                    for i, width in enumerate(column_widths):
                        col_letter = chr(65 + i)  # A, B, C...
                        worksheet.column_dimensions[col_letter].width = width
                    
                    # 设置表头样式
                    from openpyxl.styles import Font, PatternFill, Alignment
                    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    header_font = Font(bold=True, color="FFFFFF")
                    header_align = Alignment(horizontal="center")
                    
                    for cell in worksheet[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = header_align
                    
                    # 添加汇总行
                    last_row = len(bom_df) + 2
                    total_cost = sum(float(price) for price in bom_df["总价 (CNY)"])
                    worksheet[f'A{last_row}'] = "总计"
                    worksheet[f'J{last_row}'] = f"{total_cost:.2f}"
                    worksheet[f'J{last_row}'].font = Font(bold=True)
                
                logger.info(f"✓ Excel BOM 表已保存：{output_path}")
            
            else:
                # CSV 格式
                if not output_path.endswith('.csv'):
                    output_path = output_path.rsplit('.', 1)[0] + '.csv'
                bom_df.to_csv(output_path, index=False, encoding='utf-8-sig')
                logger.info(f"✓ CSV BOM 表已保存：{output_path}")
            
            return True
        except Exception as e:
            logger.error(f"保存失败：{e}")
            return False
    
    def display_bom(self, bom_df: pd.DataFrame):
        """显示 BOM 表"""
        print("\n" + "=" * 140)
        print("BOM 表 - 物料清单（Bill of Materials）")
        print("=" * 140)
        print(bom_df.to_string(index=False))
        print("=" * 140)
        
        try:
            total_cost = sum(float(price) for price in bom_df["总价 (CNY)"])
            total_items = bom_df["数量"].sum()
            print(f"📊 元器件类型数：{len(bom_df)}  |  总数量：{total_items}  |  总成本：CNY {total_cost:.2f}")
        except Exception:
            pass


# ==================== 主识别流程 ====================

class CircuitAnalyzer:
    """电路图分析主类（OCR + AI 双引擎）"""
    
    def __init__(self, qwen_api_key: str, ocr_languages: str = "ch_sim,en"):
        self.qwen_api_key = qwen_api_key
        self.ocr = OCRRecognizer(ocr_languages)
        self.ai = AIRecognizer(qwen_api_key) if qwen_api_key else None
        logger.info("CircuitAnalyzer 初始化完成")
    
    def analyze(self, image_path: str, min_confidence: int = 5) -> List[Dict[str, str]]:
        """
        分析电路图
        流程：OCR 识别 → 检查数量 → AI 补充 → 合并结果
        """
        all_components = []
        
        # 步骤 1: OCR 识别
        logger.info("=" * 50)
        logger.info("步骤 1: OCR 图像识别")
        logger.info("=" * 50)
        
        ocr_text = self.ocr.recognize_text(image_path)
        ocr_components = self.ocr.extract_components_from_text(ocr_text)
        all_components.extend(ocr_components)
        
        logger.info(f"OCR 识别结果：{len(ocr_components)} 个元器件")
        
        # 步骤 2: 判断是否需要 AI 补充
        if len(ocr_components) < min_confidence:
            logger.info(f"OCR 识别数量 ({len(ocr_components)}) 低于阈值 ({min_confidence})，启动 AI 补充...")
            
            if self.ai:
                logger.info("=" * 50)
                logger.info("步骤 2: AI 补充识别")
                logger.info("=" * 50)
                
                ai_components = self.ai.recognize_components(image_path, ocr_components)
                
                # 合并 OCR 和 AI 结果（去重）
                all_components = self._merge_results(ocr_components, ai_components)
                logger.info(f"AI 补充后总数：{len(all_components)} 个元器件")
            else:
                logger.warning("AI 识别器未初始化，使用 OCR 结果")
        else:
            logger.info(f"OCR 识别数量充足 ({len(ocr_components)})，跳过 AI 补充")
        
        # 步骤 3: 手动确认（可选）
        if len(all_components) == 0:
            logger.warning("未识别到任何元器件，进入手动输入模式")
            all_components = self._manual_input()
        
        return all_components
    
    def _merge_results(self, ocr_components: List[Dict], ai_components: List[Dict]) -> List[Dict]:
        """合并 OCR 和 AI 结果，去重"""
        merged = {}
        
        # 先添加 OCR 结果
        for comp in ocr_components:
            key = comp['reference'].upper()
            merged[key] = comp
        
        # 添加 AI 结果（不覆盖已有）
        for comp in ai_components:
            key = comp['reference'].upper()
            if key not in merged:
                merged[key] = comp
        
        return list(merged.values())
    
    def _manual_input(self) -> List[Dict[str, str]]:
        """手动输入元器件"""
        components = []
        print("\n请输入元器件 (格式：Reference|Value|Package|Type，输入'done'结束):")
        print("示例：R1|10k|0603|电阻")
        
        while True:
            user_input = input("Component: ").strip()
            if user_input.lower() == "done":
                break
            parts = user_input.split("|")
            if len(parts) >= 3:
                components.append({
                    "reference": parts[0].strip(),
                    "value": parts[1].strip(),
                    "package": parts[2].strip(),
                    "type": parts[3].strip() if len(parts) > 3 else "",
                    "source": "手动"
                })
                logger.info(f"✓ 已添加 {parts[0]}")
            else:
                logger.warning("格式错误，请使用 Reference|Value|Package|Type")
        
        return components


# ==================== 主程序 ====================

def main():
    """主入口函数"""
    print("\n" + "=" * 60)
    print("🔧 BOM Assistant - 智能物料清单生成器")
    print("流程：OCR 识别 → AI 补充 → Excel 输出")
    print("=" * 60)
    
    # 检查依赖
    print(f"\n依赖检查:")
    print(f"  - EasyOCR: {'✓' if OCR_AVAILABLE else '✗'}")
    print(f"  - Tesseract: {'✓' if TESSERACT_AVAILABLE else '✗'}")
    print(f"  - Excel 支持：{'✓' if EXCEL_AVAILABLE else '✗'}")
    
    # 创建项目文件夹
    project_folder = create_project_folder()
    if project_folder:
        os.chdir(project_folder)
        logger.info(f"工作目录：{project_folder}")
    
    # 加载配置
    config = Config.load_from_file()
    
    # 获取 API 密钥
    if not config.qwen_api_key:
        config.qwen_api_key = input("\n请输入千问 API 密钥：").strip()
        if not config.qwen_api_key:
            logger.error("API 密钥不能为空")
            return
    
    if not config.lichuang_api_key:
        config.lichuang_api_key = input("立创商城 API 密钥 (留空使用本地数据库): ").strip()
        config.save_to_file()
    
    # 选择输出格式
    if EXCEL_AVAILABLE:
        format_choice = input("\n输出格式 (1=Excel, 2=CSV, 默认=1): ").strip()
        config.output_format = "excel" if format_choice != "2" else "csv"
    else:
        config.output_format = "csv"
        logger.info("未安装 openpyxl，使用 CSV 格式")
    
    # 选择图像文件
    image_path = select_image_file()
    if not image_path:
        logger.error("未选择图像文件")
        return
    
    # 分析电路图
    analyzer = CircuitAnalyzer(config.qwen_api_key)
    components = analyzer.analyze(image_path)
    
    if not components:
        logger.error("未识别到元器件")
        return
    
    # 生成 BOM
    logger.info("=" * 50)
    logger.info("步骤 3: 生成 BOM 表")
    logger.info("=" * 50)
    
    bom_gen = BOMGenerator(config.lichuang_api_key, config.output_format)
    bom_gen.add_components(components)
    
    merge = input("\n是否合并相同元器件？(y/n, 默认=y): ").strip().lower() != 'n'
    bom_df = bom_gen.generate_bom(merge=merge)
    bom_gen.display_bom(bom_df)
    
    # 保存 BOM
    ext = ".xlsx" if config.output_format == "excel" else ".csv"
    default_name = f"BOM_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
    output_file = input(f"\n保存文件名 (默认 {default_name}): ").strip()
    if not output_file:
        output_file = select_save_file(default_name) or default_file
    else:
        if not output_file.endswith(ext):
            output_file = output_file.rsplit('.', 1)[0] + ext
    
    if bom_gen.save_bom(bom_df, output_file):
        print(f"\n✅ BOM 生成完成！")
        print(f"  📁 文件：{output_file}")
        print(f"  📊 元器件类型：{len(bom_df)}")
        print(f"  📦 总数量：{bom_df['数量'].sum()}")
        try:
            total = sum(float(p) for p in bom_df["总价 (CNY)"])
            print(f"  💰 总成本：CNY {total:.2f}")
        except:
            pass


if __name__ == "__main__":
    main()