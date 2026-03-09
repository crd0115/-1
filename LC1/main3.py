"""
BOM Assistant - 智能物料清单生成器 (GUI 版本)
流程：AI 识别（主）→ OCR 补充（辅）→ Excel 输出
"""

import os
import json
import base64
import logging
import re
import threading
from pathlib import Path
from typing import List, Dict, Any, Optional
from dataclasses import dataclass
from datetime import datetime

import pandas as pd
import dashscope
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# OCR 库
try:
    import easyocr
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    easyocr = None

try:
    import pytesseract
    from PIL import Image as PILImage
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False
    pytesseract = None

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
    output_format: str = "excel"
    
    @classmethod
    def load_from_file(cls, path: str = "config.json") -> "Config":
        config = cls()
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                config.qwen_api_key = data.get("qwen_api_key", "")
                config.lichuang_api_key = data.get("lichuang_api_key", "")
                config.output_format = data.get("output_format", "excel")
        return config
    
    def save_to_file(self, path: str = "config.json"):
        with open(path, "w", encoding="utf-8") as f:
            json.dump({
                "qwen_api_key": self.qwen_api_key,
                "lichuang_api_key": self.lichuang_api_key,
                "output_format": self.output_format
            }, f, indent=2)


# ==================== 日志设置 ====================

class TextHandler(logging.Handler):
    """日志输出到 Text 控件"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    
    def emit(self, record):
        msg = self.format(record)
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.configure(state='disabled')
        self.text_widget.see(tk.END)


def setup_logging(text_widget=None):
    logger = logging.getLogger("BOM_Assistant")
    logger.setLevel(logging.INFO)
    
    if not logger.handlers:
        if text_widget:
            handler = TextHandler(text_widget)
        else:
            handler = logging.StreamHandler()
        
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%H:%M:%S'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    
    return logger


logger = setup_logging()


# ==================== OCR 识别模块 ====================

class OCRRecognizer:
    """OCR 图像识别器"""
    
    def __init__(self, languages: str = "ch_sim,en"):
        self.languages = languages.split(",")
        self.reader = None
        self._init_ocr()
    
    def _init_ocr(self):
        if OCR_AVAILABLE and easyocr:
            try:
                self.reader = easyocr.Reader(self.languages, gpu=False)
                logger.info("✓ EasyOCR 初始化成功")
            except Exception as e:
                logger.warning(f"EasyOCR 初始化失败：{e}")
    
    def recognize_text(self, image_path: str) -> str:
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"图像文件不存在：{image_path}")
        
        text = ""
        if self.reader:
            try:
                results = self.reader.readtext(image_path)
                text = "\n".join([result[1] for result in results])
                logger.info(f"EasyOCR 提取 {len(results)} 个文本块")
            except Exception as e:
                logger.warning(f"EasyOCR 识别失败：{e}")
        
        return text
    
    def extract_components_from_text(self, text: str) -> List[Dict[str, str]]:
        components = []
        patterns = [
            r'(R\d+)[\s:]+([\d\.]+[kKmMΩΩ]?)[\s]+(\d+)',
            r'(C\d+)[\s:]+([\d\.]+[uUnNpPfF]?)\s+(\d+)',
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
        
        seen = set()
        unique_components = []
        for comp in components:
            key = f"{comp['reference']}_{comp['value']}"
            if key not in seen:
                seen.add(key)
                unique_components.append(comp)
        
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
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    
    def recognize_components(self, image_path: str, ocr_results: List[Dict] = None) -> List[Dict[str, str]]:
        try:
            image_base64 = self.load_image_base64(image_path)
            prompt = self._build_prompt(ocr_results)
            
            logger.info("正在调用 Qwen AI 进行元件识别...")
            response = dashscope.MultiModalConversation.call(
                model='qwen-vl-max',
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
        base_prompt = """你是一位专业的电子工程师，请仔细分析这张电路图。

【任务要求】
1. 识别图中所有电子元器件
2. 输出严格的 JSON 数组格式，不要任何其他文字
3. 每个元件必须包含以下字段：
   - reference：参考编号（如 R1、C2、U3、Q1）
   - value：元件值（如 10k、100nF、STM32F103C8T6）
   - package：封装（如 0603、0805、LQFP48、SOT23）
   - type：元件类型（Resistor/Capacitor/IC/Transistor/Diode/Connector）

【输出格式示例】
[{"reference": "R1", "value": "10k", "package": "0603", "type": "Resistor"}]
"""
        if ocr_results:
            ocr_text = "\n".join([f"{c['reference']} {c['value']} {c['package']}" for c in ocr_results[:15]])
            base_prompt += f"\n【OCR 参考信息】:\n{ocr_text}"
        
        return base_prompt
    
    def _parse_ai_response(self, content: str) -> List[Dict[str, Any]]:
        if isinstance(content, list):
            return self._normalize_components(content)
        
        try:
            content = re.sub(r'```json\s*|\s*```', '', content.strip())
            content = re.sub(r'```\s*|\s*```', '', content)
            
            json_match = re.search(r'\[[\s\S]*\]', content)
            if json_match:
                content = json_match.group()
            
            parsed = json.loads(content)
            
            if isinstance(parsed, list):
                return self._normalize_components(parsed)
            elif isinstance(parsed, dict) and 'components' in parsed:
                return self._normalize_components(parsed['components'])
                
        except json.JSONDecodeError as e:
            logger.warning(f"JSON 解析失败：{e}")
        
        return self._extract_from_text(content)
    
    def _normalize_components(self, components: List[Dict]) -> List[Dict[str, Any]]:
        normalized = []
        for comp in components:
            if not isinstance(comp, dict):
                continue
            
            normalized_comp = {
                'reference': comp.get('reference', comp.get('ref', comp.get('name', ''))),
                'value': comp.get('value', comp.get('val', '')),
                'package': comp.get('package', comp.get('pkg', '')),
                'type': comp.get('type', comp.get('component_type', 'Unknown')),
                'source': 'AI'
            }
            
            if normalized_comp['reference'] and normalized_comp['value']:
                normalized.append(normalized_comp)
        
        return normalized
    
    def _extract_from_text(self, text: str) -> List[Dict[str, str]]:
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
        self.component_database = self._init_component_database()
    
    def _init_component_database(self) -> Dict[str, Dict[str, Any]]:
        return {
            "0.1uF": {"part_number": "C0603X5R1V104K030BC", "name": "贴片电容 0.1uF 50V X7R 0603", "price": 0.02, "stock": 50000},
            "10uF": {"part_number": "C0805X5R1V106K030BC", "name": "贴片电容 10uF 25V X5R 0805", "price": 0.05, "stock": 20000},
            "10k": {"part_number": "R0603-10K", "name": "贴片电阻 10kΩ 1% 0603", "price": 0.005, "stock": 100000},
            "1k": {"part_number": "R0603-1K", "name": "贴片电阻 1kΩ 1% 0603", "price": 0.005, "stock": 100000},
            "STM32F103C8T6": {"part_number": "STM32F103C8T6-LQFP48", "name": "STM32F103C8T6 微控制器 LQFP48", "price": 8.5, "stock": 500},
            "LM1117-3.3": {"part_number": "LM1117-3.3-SOT223", "name": "LM1117-3.3 线性稳压芯片 SOT-223", "price": 0.5, "stock": 1000},
        }
    
    def search_component(self, component_name: str, package: str = "") -> List[Dict[str, Any]]:
        search_results = []
        for key, comp_data in self.component_database.items():
            if component_name.lower() in key.lower() or component_name.lower() in comp_data["name"].lower():
                search_results.append({
                    "part_number": comp_data["part_number"],
                    "name": comp_data["name"],
                    "price": comp_data["price"],
                    "stock": comp_data["stock"],
                    "supplier": "立创商城"
                })
        
        if not search_results:
            search_results = [{
                "part_number": f"UNKNOWN-{component_name}",
                "name": f"{component_name} (需手动确认)",
                "price": 0.0,
                "stock": 0,
                "supplier": "立创商城"
            }]
        
        return search_results


# ==================== GUI 主窗口 ====================

class BOMAssistantGUI:
    """BOM Assistant 图形界面"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("🔧 BOM Assistant - 智能物料清单生成器")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # 配置
        self.config = Config.load_from_file()
        self.image_path = None
        self.components = []
        self.bom_df = None
        
        # 识别器
        self.ocr = None
        self.ai = None
        
        # 创建界面
        self._create_menu()
        self._create_main_layout()
        self._create_status_bar()
        
        # 初始化日志
        setup_logging(self.log_text)
        
        logger.info("BOM Assistant GUI 启动完成")
    
    def _create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="打开图像", command=self._select_image, accelerator="Ctrl+O")
        file_menu.add_command(label="保存 BOM", command=self._save_bom, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        
        # 配置菜单
        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="配置", menu=config_menu)
        config_menu.add_command(label="API 设置", command=self._show_config_dialog)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self._show_help)
        help_menu.add_command(label="关于", command=self._show_about)
        
        # 快捷键绑定
        self.root.bind('<Control-o>', lambda e: self._select_image())
        self.root.bind('<Control-s>', lambda e: self._save_bom())
    
    def _create_main_layout(self):
        """创建主布局"""
        # 主容器
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 左侧面板（图像和配置）
        left_panel = ttk.LabelFrame(main_frame, text="图像预览", padding="10")
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 图像显示区域
        self.image_label = ttk.Label(left_panel, text="未加载图像\n\n请点击\"选择图像\"按钮", 
                                      anchor=tk.CENTER, background="#f0f0f0")
        self.image_label.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 图像操作按钮
        btn_frame = ttk.Frame(left_panel)
        btn_frame.pack(fill=tk.X)
        
        self.btn_select = ttk.Button(btn_frame, text="📁 选择图像", command=self._select_image)
        self.btn_select.pack(side=tk.LEFT, padx=(0, 5))
        
        self.btn_clear = ttk.Button(btn_frame, text="🗑️ 清除", command=self._clear_image)
        self.btn_clear.pack(side=tk.LEFT)
        
        # 右侧面板（识别和结果）
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 配置区域
        config_frame = ttk.LabelFrame(right_panel, text="配置选项", padding="10")
        config_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(config_frame, text="识别模式:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.mode_var = tk.StringVar(value="ai_primary")
        mode_combo = ttk.Combobox(config_frame, textvariable=self.mode_var, width=20, state="readonly")
        mode_combo['values'] = ('AI 优先', 'OCR 优先', '仅 AI', '仅 OCR')
        mode_combo.grid(row=0, column=1, sticky=tk.W)
        
        ttk.Label(config_frame, text="输出格式:").grid(row=0, column=2, sticky=tk.W, padx=(20, 10))
        self.format_var = tk.StringVar(value=self.config.output_format)
        format_combo = ttk.Combobox(config_frame, textvariable=self.format_var, width=10, state="readonly")
        format_combo['values'] = ('excel', 'csv')
        format_combo.grid(row=0, column=3, sticky=tk.W)
        
        # 识别按钮
        self.btn_recognize = ttk.Button(config_frame, text="🚀 开始识别", command=self._start_recognition)
        self.btn_recognize.grid(row=0, column=4, padx=(20, 0))
        
        # BOM 表格区域
        table_frame = ttk.LabelFrame(right_panel, text="BOM 表", padding="10")
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 创建 Treeview
        columns = ("reference", "value", "package", "type", "quantity", "part_number", "price", "source")
        self.bom_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        
        # 设置列标题
        headers = {
            "reference": "参考编号",
            "value": "值/型号",
            "package": "封装",
            "type": "类型",
            "quantity": "数量",
            "part_number": "立创编号",
            "price": "单价 (CNY)",
            "source": "来源"
        }
        
        for col in columns:
            self.bom_tree.heading(col, text=headers[col])
            self.bom_tree.column(col, width=100)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.bom_tree.yview)
        self.bom_tree.configure(yscrollcommand=scrollbar.set)
        
        self.bom_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 统计信息
        self.stats_label = ttk.Label(table_frame, text="", foreground="blue")
        self.stats_label.pack(fill=tk.X, pady=(5, 0))
        
        # 日志区域
        log_frame = ttk.LabelFrame(right_panel, text="运行日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)
    
    def _create_status_bar(self):
        """创建状态栏"""
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def _select_image(self):
        """选择图像文件"""
        filetypes = [
            ("图像文件", "*.png *.jpg *.jpeg *.bmp *.tiff"),
            ("所有文件", "*.*")
        ]
        
        path = filedialog.askopenfilename(
            title="选择电路图文件",
            filetypes=filetypes
        )
        
        if path:
            self._load_image(path)
    
    def _load_image(self, path):
        """加载并显示图像"""
        try:
            self.image_path = path
            
            # 加载图像
            img = Image.open(path)
            
            # 缩放图像以适应显示区域
            max_width = 400
            max_height = 400
            ratio = min(max_width / img.width, max_height / img.height)
            new_size = (int(img.width * ratio), int(img.height * ratio))
            img = img.resize(new_size, Image.Resampling.LANCZOS)
            
            # 转换为 PhotoImage
            self.photo = ImageTk.PhotoImage(img)
            self.image_label.config(image=self.photo, text="")
            
            logger.info(f"已加载图像：{os.path.basename(path)}")
            self.status_var.set(f"图像：{os.path.basename(path)}")
            
        except Exception as e:
            messagebox.showerror("错误", f"加载图像失败：{e}")
            logger.error(f"加载图像失败：{e}")
    
    def _clear_image(self):
        """清除图像"""
        self.image_path = None
        self.image_label.config(image="", text="未加载图像\n\n请点击\"选择图像\"按钮")
        self.components = []
        self.bom_df = None
        self._clear_bom_table()
        logger.info("已清除图像")
        self.status_var.set("就绪")
    
    def _clear_bom_table(self):
        """清空 BOM 表格"""
        for item in self.bom_tree.get_children():
            self.bom_tree.delete(item)
        self.stats_label.config(text="")
    
    def _start_recognition(self):
        """开始识别（多线程）"""
        if not self.image_path:
            messagebox.showwarning("警告", "请先选择图像文件！")
            return
        
        if not self.config.qwen_api_key:
            messagebox.showwarning("警告", "请先配置千问 API 密钥！")
            self._show_config_dialog()
            return
        
        # 禁用按钮
        self.btn_recognize.config(state='disabled')
        self.status_var.set("正在识别...")
        
        # 启动识别线程
        thread = threading.Thread(target=self._recognition_thread)
        thread.daemon = True
        thread.start()
    
    def _recognition_thread(self):
        """识别线程"""
        try:
            # 初始化识别器
            if not self.ai:
                self.ai = AIRecognizer(self.config.qwen_api_key)
            if not self.ocr:
                self.ocr = OCRRecognizer()
            
            # 清空表格
            self.root.after(0, self._clear_bom_table)
            
            # 执行识别
            mode = self.mode_var.get()
            
            ai_components = []
            ocr_components = []
            
            # AI 识别
            if mode in ["AI 优先", "仅 AI"]:
                logger.info("开始 AI 识别...")
                ai_components = self.ai.recognize_components(self.image_path)
                logger.info(f"AI 识别完成：{len(ai_components)} 个元器件")
            
            # OCR 识别
            if mode in ["OCR 优先", "仅 OCR"]:
                logger.info("开始 OCR 识别...")
                ocr_text = self.ocr.recognize_text(self.image_path)
                ocr_components = self.ocr.extract_components_from_text(ocr_text)
                logger.info(f"OCR 识别完成：{len(ocr_components)} 个元器件")
            
            # 合并结果
            if mode == "AI 优先":
                self.components = self._merge_results(ai_components, ocr_components)
            elif mode == "OCR 优先":
                self.components = self._merge_results(ocr_components, ai_components)
            else:
                self.components = ai_components + ocr_components
            
            # 丰富信息
            self._enrich_components()
            
            # 生成 BOM
            self._generate_bom()
            
            # 更新 UI
            self.root.after(0, self._recognition_complete)
            
        except Exception as e:
            logger.error(f"识别失败：{e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"识别失败：{e}"))
            self.root.after(0, self._recognition_error)
    
    def _merge_results(self, primary: List[Dict], secondary: List[Dict]) -> List[Dict]:
        """合并结果"""
        merged = {}
        for comp in primary:
            key = comp['reference'].upper()
            merged[key] = comp
        for comp in secondary:
            key = comp['reference'].upper()
            if key not in merged:
                merged[key] = comp
        return list(merged.values())
    
    def _enrich_components(self):
        """丰富元器件信息"""
        lichuang = LiChuangAPI(self.config.lichuang_api_key)
        
        for comp in self.components:
            search_results = lichuang.search_component(comp.get("value", ""), comp.get("package", ""))
            if search_results:
                best = search_results[0]
                comp["part_number"] = best.get("part_number", "")
                comp["price"] = best.get("price", 0)
                comp["stock"] = best.get("stock", 0)
                comp["supplier"] = best.get("supplier", "立创商城")
    
    def _generate_bom(self):
        """生成 BOM 表"""
        if not self.components:
            return
        
        # 合并相同元器件
        merged = {}
        for comp in self.components:
            key = f"{comp['value']}_{comp['package']}"
            if key not in merged:
                merged[key] = {
                    "value": comp["value"],
                    "package": comp["package"],
                    "references": [comp["reference"]],
                    "quantity": 1,
                    "type": comp.get("type", ""),
                    "part_number": comp.get("part_number", ""),
                    "price": comp.get("price", 0),
                    "source": comp.get("source", "")
                }
            else:
                merged[key]["references"].append(comp["reference"])
                merged[key]["quantity"] += 1
        
        # 创建 DataFrame
        data = []
        for comp in merged.values():
            data.append({
                "reference": ",".join(comp["references"]),
                "value": comp["value"],
                "package": comp["package"],
                "type": comp["type"],
                "quantity": comp["quantity"],
                "part_number": comp["part_number"],
                "price": f"{comp['price']:.4f}",
                "source": comp["source"]
            })
        
        self.bom_df = pd.DataFrame(data)
    
    def _recognition_complete(self):
        """识别完成"""
        self.btn_recognize.config(state='normal')
        
        if self.bom_df is not None and len(self.bom_df) > 0:
            # 填充表格
            for _, row in self.bom_df.iterrows():
                self.bom_tree.insert("", tk.END, values=list(row))
            
            # 更新统计
            total_items = self.bom_df["quantity"].sum()
            total_cost = sum(float(p) * int(q) for p, q in zip(self.bom_df["price"], self.bom_df["quantity"]))
            self.stats_label.config(text=f"元器件类型：{len(self.bom_df)}  |  总数量：{total_items}  |  总成本：CNY {total_cost:.2f}")
            
            self.status_var.set(f"识别完成 - {len(self.bom_df)} 个元器件类型")
            logger.info(f"识别完成！共 {len(self.bom_df)} 个元器件类型")
        else:
            self.status_var.set("未识别到元器件")
            logger.warning("未识别到元器件")
    
    def _recognition_error(self):
        """识别错误"""
        self.btn_recognize.config(state='normal')
        self.status_var.set("识别失败")
    
    def _save_bom(self):
        """保存 BOM 表"""
        if self.bom_df is None or len(self.bom_df) == 0:
            messagebox.showwarning("警告", "没有可保存的 BOM 数据！")
            return
        
        ext = ".xlsx" if self.format_var.get() == "excel" else ".csv"
        default_name = f"BOM_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
        
        filetypes = [("Excel 文件", "*.xlsx"), ("CSV 文件", "*.csv")] if self.format_var.get() == "excel" else [("CSV 文件", "*.csv")]
        
        path = filedialog.asksaveasfilename(
            title="保存 BOM 表",
            defaultextension=ext,
            initialfile=default_name,
            filetypes=filetypes
        )
        
        if path:
            try:
                if path.endswith(".xlsx"):
                    self.bom_df.to_excel(path, index=False)
                else:
                    self.bom_df.to_csv(path, index=False, encoding='utf-8-sig')
                
                messagebox.showinfo("成功", f"BOM 表已保存：\n{path}")
                logger.info(f"BOM 表已保存：{path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败：{e}")
                logger.error(f"保存失败：{e}")
    
    def _show_config_dialog(self):
        """显示配置对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title("配置")
        dialog.geometry("500x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # API 密钥
        ttk.Label(frame, text="千问 API 密钥:").grid(row=0, column=0, sticky=tk.W, pady=5)
        api_entry = ttk.Entry(frame, width=50)
        api_entry.grid(row=0, column=1, pady=5, padx=(10, 0))
        api_entry.insert(0, self.config.qwen_api_key)
        
        ttk.Label(frame, text="立创 API 密钥:").grid(row=1, column=0, sticky=tk.W, pady=5)
        lc_entry = ttk.Entry(frame, width=50)
        lc_entry.grid(row=1, column=1, pady=5, padx=(10, 0))
        lc_entry.insert(0, self.config.lichuang_api_key)
        
        # 保存按钮
        def save_config():
            self.config.qwen_api_key = api_entry.get().strip()
            self.config.lichuang_api_key = lc_entry.get().strip()
            self.config.output_format = self.format_var.get()
            self.config.save_to_file()
            
            # 重新初始化 AI
            if self.config.qwen_api_key:
                try:
                    self.ai = AIRecognizer(self.config.qwen_api_key)
                except Exception as e:
                    messagebox.showerror("错误", f"AI 初始化失败：{e}")
            
            messagebox.showinfo("成功", "配置已保存！")
            dialog.destroy()
        
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=20)
        
        ttk.Button(btn_frame, text="保存", command=save_config).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT)
    
    def _show_help(self):
        """显示帮助"""
        help_text = """
BOM Assistant 使用说明

1. 配置 API 密钥
   - 点击"配置" → "API 设置"
   - 输入千问 API 密钥（必需）
   - 输入立创 API 密钥（可选）

2. 加载电路图
   - 点击"选择图像"按钮
   - 支持 PNG、JPG、BMP 等格式

3. 开始识别
   - 选择识别模式（推荐"AI 优先"）
   - 点击"开始识别"按钮
   - 等待识别完成

4. 保存结果
   - 查看 BOM 表格
   - 点击"文件" → "保存 BOM"
   - 选择 Excel 或 CSV 格式

提示：
- AI 识别需要有效的 API 密钥
- 清晰的图像可获得更好的识别效果
- 识别结果建议人工核对
"""
        messagebox.showinfo("使用说明", help_text)
    
    def _show_about(self):
        """显示关于"""
        about_text = """
BOM Assistant v1.0

智能物料清单生成器

功能：
- AI 图像识别元器件
- OCR 文字识别
- 立创商城数据对接
- Excel/CSV 导出

技术栈：
- Python + Tkinter
- 通义千问 VL 模型
- EasyOCR

© 2024 All Rights Reserved
"""
        messagebox.showinfo("关于", about_text)


# ==================== 主程序 ====================

def main():
    root = tk.Tk()
    
    # 设置样式
    style = ttk.Style()
    style.theme_use('clam')
    
    # 创建应用
    app = BOMAssistantGUI(root)
    
    # 运行
    root.mainloop()


if __name__ == "__main__":
    main()