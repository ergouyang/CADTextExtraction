import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import ezdxf
import csv

class DWGProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DWG 文字块提取工具")
        
        # ODA 转换器路径
        self.oda_path = r".\oda\ODAFileConverter.exe"
        
        # 初始化变量
        self.target_folder = ""
        self.selected_text_pattern = None
        self.output_data = []
        self.text_entities = []
        
        # 创建临时目录
        self.temp_input = os.path.abspath("temp_input")
        self.temp_output = os.path.abspath("temp_output")
        self.temp_batch_output = os.path.abspath("temp_batch_output")
        os.makedirs(self.temp_input, exist_ok=True)
        os.makedirs(self.temp_output, exist_ok=True)
        os.makedirs(self.temp_batch_output, exist_ok=True)
        
        # 创建界面布局
        self.create_widgets()
    
    def create_widgets(self):
        # 左侧面板：文件列表
        left_frame = ttk.Frame(self.root)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=5, pady=5)
        
        ttk.Label(left_frame, text="DWG文件列表").pack()
        self.file_list = tk.Listbox(left_frame)
        self.file_list.pack(expand=True, fill=tk.BOTH)
        self.file_list.bind("<<ListboxSelect>>", self.on_file_select)
        
        ttk.Button(left_frame, text="选择文件夹", command=self.load_folder).pack(pady=5)
        
        # 右侧面板：文字块选择和操作
        right_frame = ttk.Frame(self.root)
        right_frame.pack(side=tk.RIGHT, expand=True, fill=tk.BOTH, padx=5, pady=5)
        
        ttk.Label(right_frame, text="检测到的文字块").pack()
        self.text_blocks_list = tk.Listbox(right_frame)
        self.text_blocks_list.pack(expand=True, fill=tk.BOTH)
        
        ttk.Button(right_frame, text="设为特征文字", command=self.select_text_pattern).pack(pady=5)
        ttk.Button(right_frame, text="批量提取", command=self.batch_process).pack(pady=5)
        ttk.Button(right_frame, text="导出结果", command=self.export_results).pack(pady=5)
    
    def load_folder(self):
        self.target_folder = filedialog.askdirectory()
        if not self.target_folder:
            return
        
        self.file_list.delete(0, tk.END)
        for f in os.listdir(self.target_folder):
            if f.lower().endswith(".dwg"):
                self.file_list.insert(tk.END, f)
    
    def on_file_select(self, event):
        selected = self.file_list.curselection()
        if not selected:
            return
        
        # 清空临时目录
        for folder in [self.temp_input, self.temp_output]:
            for f in os.listdir(folder):
                os.remove(os.path.join(folder, f))
        
        # 复制选中文件到输入目录
        original_file = os.path.join(self.target_folder, self.file_list.get(selected[0]))
        temp_dwg = os.path.join(self.temp_input, os.path.basename(original_file))
        
        try:
            import shutil
            shutil.copy(original_file, temp_dwg)
        except Exception as e:
            messagebox.showerror("错误", f"文件准备失败: {str(e)}")
            return

        # ODA转换命令
        try:
            command = [
                self.oda_path,
                f'{self.temp_input}',
                f'{self.temp_output}',
                "ACAD2018",
                "DXF",
                "0",
                "1",
                "*.dwg"
            ]
            subprocess.run(command, check=True, 
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                shell=True,
                timeout=60,
                cwd=os.path.dirname(self.oda_path))
            
            # 获取生成的DXF路径
            base_name = os.path.splitext(os.path.basename(temp_dwg))[0]
            dxf_file = os.path.join(self.temp_output, f"{base_name}.dxf")
            
            # 删除temp_input中的DWG文件
            if os.path.exists(temp_dwg):
                os.remove(temp_dwg)
            
            if not os.path.exists(dxf_file):
                raise Exception("DXF文件生成失败")
            
            # 提取文字块
            doc = ezdxf.readfile(dxf_file)
            self.extract_text_blocks(doc)
            os.remove(dxf_file)
            
        except Exception as e:
            messagebox.showerror("错误", f"转换失败: {str(e)}")
            return
    
    # 修改文字块信息记录方式（在 extract_text_blocks 方法中）
    def extract_text_blocks(self, doc):
        self.text_blocks_list.delete(0, tk.END)
        self.text_entities = []
        
        def extract_from_entity(entity):
            if entity.dxftype() == "TEXT":
                return [{
                    "text": entity.dxf.text,
                    "pos": (entity.dxf.insert.x, entity.dxf.insert.y)  # 记录坐标
                }]
            elif entity.dxftype() == "INSERT":
                texts = []
                for attrib in entity.attribs:
                    if attrib.dxftype() == "ATTRIB":
                        texts.append({
                            "text": attrib.dxf.text,
                            "pos": (attrib.dxf.insert.x, attrib.dxf.insert.y)  # 记录坐标
                        })
                return texts
            return []

        msp = doc.modelspace()
        for entity in msp:
            for text_info in extract_from_entity(entity):
                display_text = f"{entity.dxf.layer}: {text_info['text']}"
                self.text_blocks_list.insert(tk.END, display_text)
                self.text_entities.append({
                    "text": text_info['text'],
                    "layer": entity.dxf.layer,
                    "pos": text_info['pos']  # 存储坐标信息
                })
    

    # 修改特征选择方法
    def select_text_pattern(self):
        selected = self.text_blocks_list.curselection()
        if selected:
            selected_info = self.text_entities[selected[0]]
            # 存储位置信息和阈值
            self.selected_text_pattern = {
                "layer": selected_info["layer"],
                "position": selected_info["pos"],  # 存储坐标
                "threshold": 20.0  # 单位距离阈值
            }
            messagebox.showinfo("特征设置", f"已设置位置特征: {selected_info['pos']}")
    # 修改批量处理匹配逻辑
    def batch_process(self):
        if not self.selected_text_pattern:
            messagebox.showwarning("警告", "请先选择特征文字块位置")
            return
        
        self.output_data = []
        # 批量转换DWG到DXF
        subprocess.run([
            self.oda_path,
            self.target_folder,
            self.temp_batch_output,
            "ACAD2018", "DXF", "0", "1"
        ], check=True)

        for dxf_file in os.listdir(self.temp_batch_output):
            dxf_temp = os.path.join(self.temp_batch_output, dxf_file)



        # for dwg_file in os.listdir(self.target_folder):
        #     if not dwg_file.lower().endswith(".dwg"):
        #         continue
        #
        #     full_path = os.path.join(self.target_folder, dwg_file)
        #     dxf_temp = "temp_batch.dxf"
        #
        #     try:
        #         # 转换 DWG 到 DXF
        #         subprocess.run([
        #             self.oda_path,
        #             full_path,
        #             dxf_temp,
        #             "ACAD2018", "DXF", "0", "1"
        #         ], check=True)
            try:
                # 提取匹配的文字
                doc = ezdxf.readfile(dxf_temp)
                matched_texts = []

                # 遍历实体检查位置
                for entity in doc.modelspace():
                    texts = []
                    if entity.dxftype() == "TEXT":
                        texts.append({
                            "text": entity.dxf.text,
                            "pos": (entity.dxf.insert.x, entity.dxf.insert.y),
                            "layer": entity.dxf.layer
                        })
                    elif entity.dxftype() == "INSERT":
                        for attrib in entity.attribs:
                            if attrib.dxftype() == "ATTRIB":
                                texts.append({
                                    "text": attrib.dxf.text,
                                    "pos": (attrib.dxf.insert.x, attrib.dxf.insert.y),
                                    "layer": entity.dxf.layer
                                })

                    for text in texts:
                        # 计算欧氏距离
                        dx = text["pos"][0] - self.selected_text_pattern["position"][0]
                        dy = text["pos"][1] - self.selected_text_pattern["position"][1]
                        distance = (dx**2 + dy**2)**0.5

                        if (distance <= self.selected_text_pattern["threshold"] and
                            text["layer"] == self.selected_text_pattern["layer"]):
                            matched_texts.append(text["text"])

                if matched_texts:
                    self.output_data.append({
                        "filename": dxf_file,
                        "text": ", ".join(matched_texts)
                    })

                os.remove(dxf_temp)
            except Exception as e:
                print(f"处理文件 {dxf_file} 时出错: {str(e)}")
                continue
    
    def export_results(self):
        if not self.output_data:
            return
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv")]
        )
        if not save_path:
            return
        
        with open(save_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["文件名", "提取文字"])
            for item in self.output_data:
                writer.writerow([item["filename"], item["text"]])

if __name__ == "__main__":
    root = tk.Tk()
    app = DWGProcessorApp(root)
    root.mainloop()
