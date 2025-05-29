import json
import os
import shutil
from pathlib import Path
from typing import Dict, List, Any
import zipfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

class AivisSpeechStyleAdder:
    """AivisSpeechモデルに複数のスタイルを追加するクラス"""
    
    def __init__(self):
        self.styles = {
            "normal": "ノーマル",
            "standard": "通常",
            "high_tension": "テンション高め",
            "calm": "落ち着き",
            "cheerful": "上機嫌",
            "emotional": "怒り・悲しみ"
        }
        
    def load_model(self, model_path: str) -> Dict[str, Any]:
        """AivisSpeechモデルを読み込む"""
        if model_path.endswith('.aivis'):
            return self._load_aivis_model(model_path)
        elif model_path.endswith('.json'):
            with open(model_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            raise ValueError("サポートされていないファイル形式です")
    
    def _load_aivis_model(self, model_path: str) -> Dict[str, Any]:
        """ZIP形式のAivisモデルを読み込む"""
        temp_dir = Path("temp_aivis_extract")
        temp_dir.mkdir(exist_ok=True)
        
        try:
            with zipfile.ZipFile(model_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            config_path = temp_dir / "config.json"
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                return {"type": "aivis", "config": config, "temp_dir": str(temp_dir)}
            else:
                raise FileNotFoundError("config.jsonが見つかりません")
        except Exception as e:
            shutil.rmtree(temp_dir, ignore_errors=True)
            raise e
    
    def add_styles_to_model(self, model_data: Dict[str, Any]) -> Dict[str, Any]:
        """モデルに複数のスタイルを追加"""
        if model_data.get("type") == "aivis":
            config = model_data["config"]
        else:
            config = model_data
        
        if "styles" not in config:
            config["styles"] = {}
        
        for style_id, style_name in self.styles.items():
            config["styles"][style_id] = {
                "name": style_name,
                "parameters": self._get_style_parameters(style_id)
            }
        
        if "speakers" in config:
            for speaker in config["speakers"]:
                if "styles" not in speaker:
                    speaker["styles"] = list(self.styles.keys())
        
        return model_data
    
    def _get_style_parameters(self, style_id: str) -> Dict[str, float]:
        """各スタイルのパラメータを取得"""
        style_params = {
            "normal": {
                "speed": 1.0,
                "pitch": 0.0,
                "intonation": 1.0,
                "volume": 1.0,
                "emotion_strength": 0.5
            },
            "standard": {
                "speed": 1.0,
                "pitch": 0.0,
                "intonation": 1.0,
                "volume": 1.0,
                "emotion_strength": 0.5
            },
            "high_tension": {
                "speed": 1.2,
                "pitch": 0.2,
                "intonation": 1.5,
                "volume": 1.2,
                "emotion_strength": 0.8
            },
            "calm": {
                "speed": 0.9,
                "pitch": -0.1,
                "intonation": 0.8,
                "volume": 0.9,
                "emotion_strength": 0.3
            },
            "cheerful": {
                "speed": 1.1,
                "pitch": 0.1,
                "intonation": 1.3,
                "volume": 1.1,
                "emotion_strength": 0.7
            },
            "emotional": {
                "speed": 0.95,
                "pitch": -0.2,
                "intonation": 1.4,
                "volume": 1.0,
                "emotion_strength": 0.9
            }
        }
        
        return style_params.get(style_id, style_params["normal"])
    
    def save_model(self, model_data: Dict[str, Any], output_path: str):
        """スタイルが追加されたモデルを保存"""
        if model_data.get("type") == "aivis":
            self._save_aivis_model(model_data, output_path)
        else:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(model_data, f, ensure_ascii=False, indent=2)
    
    def _save_aivis_model(self, model_data: Dict[str, Any], output_path: str):
        """ZIP形式でAivisモデルを保存"""
        temp_dir = Path(model_data["temp_dir"])
        
        config_path = temp_dir / "config.json"
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(model_data["config"], f, ensure_ascii=False, indent=2)
        
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in temp_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(temp_dir)
                    zipf.write(file_path, arcname)
        
        shutil.rmtree(temp_dir, ignore_errors=True)


class AivisStyleGUI:
    """GUI インターフェース"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("AivisSpeech スタイル追加ツール")
        self.root.geometry("700x600")
        
        # スタイルを設定
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.adder = AivisSpeechStyleAdder()
        self.input_files = []
        self.output_dir = ""
        
        self.setup_ui()
        
    def setup_ui(self):
        """UIを構築"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # タイトル
        title_label = ttk.Label(main_frame, text="AivisSpeech スタイル追加ツール", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # 入力ファイル選択
        ttk.Label(main_frame, text="モデルファイル選択:", font=('Arial', 12)).grid(
            row=1, column=0, sticky=tk.W, pady=10)
        
        # ファイル選択ボタン
        select_btn = ttk.Button(main_frame, text="ファイルを選択", 
                               command=self.select_files, width=20)
        select_btn.grid(row=1, column=1, padx=10)
        
        clear_btn = ttk.Button(main_frame, text="クリア", 
                              command=self.clear_files, width=10)
        clear_btn.grid(row=1, column=2)
        
        # ファイルリスト
        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(list_frame, height=8, 
                                       yscrollcommand=scrollbar.set,
                                       font=('Arial', 10))
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # 出力ディレクトリ選択
        ttk.Label(main_frame, text="出力フォルダ:", font=('Arial', 12)).grid(
            row=3, column=0, sticky=tk.W, pady=10)
        
        output_btn = ttk.Button(main_frame, text="フォルダを選択", 
                               command=self.select_output_dir, width=20)
        output_btn.grid(row=3, column=1, padx=10)
        
        self.output_label = ttk.Label(main_frame, text="未選択", 
                                     font=('Arial', 10), foreground='gray')
        self.output_label.grid(row=4, column=0, columnspan=3, pady=5)
        
        # スタイル情報
        style_frame = ttk.LabelFrame(main_frame, text="追加されるスタイル", padding="10")
        style_frame.grid(row=5, column=0, columnspan=3, pady=20, sticky=(tk.W, tk.E))
        
        styles_text = "• ノーマル\n• 通常\n• テンション高め\n• 落ち着き\n• 上機嫌\n• 怒り・悲しみ"
        ttk.Label(style_frame, text=styles_text, font=('Arial', 11)).pack()
        
        # プログレスバー
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))
        
        # ステータスラベル
        self.status_label = ttk.Label(main_frame, text="ファイルを選択してください", 
                                     font=('Arial', 10))
        self.status_label.grid(row=7, column=0, columnspan=3, pady=5)
        
        # 実行ボタン
        self.execute_btn = ttk.Button(main_frame, text="スタイル追加を実行", 
                                     command=self.execute_process,
                                     style='Accent.TButton')
        self.execute_btn.grid(row=8, column=0, columnspan=3, pady=20)
        
        # グリッドの重み設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
    def select_files(self):
        """ファイル選択ダイアログ"""
        files = filedialog.askopenfilenames(
            title="AivisSpeechモデルを選択",
            filetypes=[("AivisSpeech Models", "*.aivis *.json"), 
                      ("All Files", "*.*")]
        )
        
        if files:
            self.input_files.extend(files)
            self.update_file_list()
            self.status_label.config(text=f"{len(self.input_files)}個のファイルが選択されました")
    
    def clear_files(self):
        """ファイルリストをクリア"""
        self.input_files = []
        self.file_listbox.delete(0, tk.END)
        self.status_label.config(text="ファイルを選択してください")
    
    def update_file_list(self):
        """ファイルリストを更新"""
        self.file_listbox.delete(0, tk.END)
        for file in self.input_files:
            self.file_listbox.insert(tk.END, os.path.basename(file))
    
    def select_output_dir(self):
        """出力ディレクトリ選択"""
        directory = filedialog.askdirectory(title="出力フォルダを選択")
        if directory:
            self.output_dir = directory
            self.output_label.config(text=f"出力先: {directory}", foreground='black')
    
    def execute_process(self):
        """処理を実行"""
        if not self.input_files:
            messagebox.showwarning("警告", "モデルファイルを選択してください")
            return
        
        if not self.output_dir:
            messagebox.showwarning("警告", "出力フォルダを選択してください")
            return
        
        # 別スレッドで処理を実行
        self.execute_btn.config(state='disabled')
        self.progress.start()
        
        thread = threading.Thread(target=self.process_files)
        thread.start()
    
    def process_files(self):
        """ファイルを処理"""
        success_count = 0
        error_count = 0
        
        for i, input_file in enumerate(self.input_files):
            try:
                # ステータス更新
                self.root.after(0, self.update_status, 
                              f"処理中... ({i+1}/{len(self.input_files)})")
                
                # 出力ファイル名を生成
                base_name = os.path.basename(input_file)
                name_without_ext = os.path.splitext(base_name)[0]
                ext = os.path.splitext(base_name)[1]
                output_file = os.path.join(self.output_dir, 
                                         f"{name_without_ext}_styled{ext}")
                
                # モデルを処理
                model_data = self.adder.load_model(input_file)
                model_data = self.adder.add_styles_to_model(model_data)
                self.adder.save_model(model_data, output_file)
                
                success_count += 1
                
            except Exception as e:
                error_count += 1
                self.root.after(0, messagebox.showerror, 
                              "エラー", f"{base_name}の処理中にエラー: {str(e)}")
        
        # 処理完了
        self.root.after(0, self.process_complete, success_count, error_count)
    
    def update_status(self, text):
        """ステータスを更新"""
        self.status_label.config(text=text)
    
    def process_complete(self, success_count, error_count):
        """処理完了時の処理"""
        self.progress.stop()
        self.execute_btn.config(state='normal')
        
        if error_count == 0:
            messagebox.showinfo("完了", 
                              f"すべてのファイルの処理が完了しました！\n"
                              f"成功: {success_count}個")
        else:
            messagebox.showwarning("完了", 
                                 f"処理が完了しました。\n"
                                 f"成功: {success_count}個\n"
                                 f"エラー: {error_count}個")
        
        self.status_label.config(text="処理完了")
    
    def run(self):
        """アプリケーションを実行"""
        self.root.mainloop()


if __name__ == "__main__":
    app = AivisStyleGUI()
    app.run()