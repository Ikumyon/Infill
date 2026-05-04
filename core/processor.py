import csv
import re
import os
from pathlib import Path
from typing import List, Dict, Optional, Any, Callable, Tuple

class CSVProcessor:
    """CSVデータの読み込みと保持を担当するクラス"""
    
    def __init__(self) -> None:
        self.headers: List[str] = []
        self.data: List[Dict[str, Any]] = []
        self.encoding: str = 'utf-8'

    def load(self, file_path: str) -> bool:
        """
        CSVファイルを読み込む。UTF-8 (BOM対応) と Shift_JIS を試行する。
        
        Args:
            file_path: 読み込むCSVファイルのパス
            
        Returns:
            読み込み成功時はTrue、失敗時はFalse
        """
        self.headers = []
        self.data = []
        encodings = ['utf-8-sig', 'shift_jis']
        for enc in encodings:
            try:
                with Path(file_path).open('r', encoding=enc, newline='') as f:
                    reader = csv.DictReader(f)
                    fieldnames = reader.fieldnames
                    if fieldnames:
                        self.headers = list(fieldnames)
                        self.data = list(reader)
                        self.encoding = enc
                        return True
            except Exception:
                continue
        return False

class TemplateProcessor:
    """ひな形テキストの解析と置換を担当するクラス"""
    
    def __init__(self) -> None:
        self.content: str = ""
        self.placeholders: List[str] = []
        self.start_delim: str = "{"
        self.end_delim: str = "}"
        self.extension: str = ".txt"

    @staticmethod
    def normalize_newlines(text: str) -> str:
        """改行コードを \n に統一する。"""
        if not text:
            return ""
        return text.replace('\r\n', '\n').replace('\r', '\n')

    def load(self, file_path: str, start_delim: str = "{", end_delim: str = "}") -> bool:
        """
        ひな形ファイルを読み込み、プレースホルダーを抽出する。
        
        Args:
            file_path: ひな形ファイルのパス
            start_delim: 開始デリミタ
            end_delim: 終了デリミタ
            
        Returns:
            成功時はTrue
        """
        self.start_delim = start_delim
        self.end_delim = end_delim
        self.content = ""
        self.placeholders = []

        if not self.start_delim or not self.end_delim:
            return False

        try:
            path = Path(file_path)
            self.extension = path.suffix if path.suffix else ".txt"
            
            with path.open('r', encoding='utf-8') as f:
                raw_content = f.read()
            
            self.content = self.normalize_newlines(raw_content)
            self.parse_placeholders()
            return True
        except Exception:
            return False

    def parse_placeholders(self) -> None:
        """現在のデリミタ設定に基づいてプレースホルダーを抽出する。"""
        s = re.escape(self.start_delim)
        e = re.escape(self.end_delim)
        pattern = rf'{s}\s*(.+?)\s*{e}'
        raw_matches = re.findall(pattern, self.content)

        seen = set()
        self.placeholders = []
        for p in raw_matches:
            p = p.strip()
            if p and p not in seen:
                seen.add(p)
                self.placeholders.append(p)

    def validate(self, csv_headers: List[str]) -> List[str]:
        """ひな形内の項目がCSV列にすべて存在するか確認し、不足している項目を返す。"""
        return [p for p in self.placeholders if p not in csv_headers]

    def preview(self, row_data: Dict[str, Any], highlight: bool = False) -> str:
        """
        指定された行データで差し込み結果を生成する。
        
        Args:
            row_data: 置換に使用するデータ
            highlight: UI表示用にHTMLハイライトを行うか
            
        Returns:
            置換後のテキスト
        """
        import html
        result = self.content
        if highlight:
            result = html.escape(result)

        s = re.escape(self.start_delim)
        e = re.escape(self.end_delim)
        
        # 1. 存在する項目を置換
        for key, value in row_data.items():
            safe_value = str(value) if value is not None else ""
            if highlight:
                display_value = f'<span style="background-color: #fff59d; color: #000; padding: 0 2px; border-radius: 2px;">{html.escape(safe_value)}</span>'
            else:
                display_value = safe_value

            pattern = rf'{s}\s*{re.escape(key)}\s*{e}'
            result = re.sub(pattern, display_value, result)
        
        # 2. 未置換の項目を赤文字にする (ハイライト時のみ)
        if highlight:
            pattern = rf'{s}\s*(.+?)\s*{e}'
            result = re.sub(pattern, r'<span style="color: #f56c6c; font-weight: bold;">\g<0></span>', result)
            result = result.replace('\n', '<br>')
            
        return result

class Exporter:
    """テキスト出力とファイル操作を担当するクラス"""

    @staticmethod
    def sanitize_filename(filename: str) -> str:
        """ファイル名に使えない文字を _ に置換する。"""
        filename = re.sub(r'[\\/:*?"<>|]', '_', filename)
        filename = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', filename)
        return filename.strip(' .')

    @staticmethod
    def get_unique_path(path: Path) -> Path:
        """ファイルが既に存在する場合、連番を振った新しいパスを返す。"""
        if not path.exists():
            return path
        
        base = path.stem
        ext = path.suffix
        parent = path.parent
        counter = 2
        while True:
            new_path = parent / f"{base}_{counter}{ext}"
            if not new_path.exists():
                return new_path
            counter += 1

    @staticmethod
    def export(content: str, output_path: str, encoding: str = 'UTF-8', newline: str = '\n') -> None:
        """
        テキストをファイルに出力する。
        
        Args:
            content: 出力内容
            output_path: 出力先パス
            encoding: 'UTF-8', 'UTF-8 BOM', 'Shift_JIS' のいずれか
            newline: 改行コード ('\\n' または '\\r\\n')
        """
        enc_map = {
            'UTF-8 BOM': 'utf-8-sig',
            'Shift_JIS': 'shift_jis'
        }
        final_encoding = enc_map.get(encoding, 'utf-8')

        # 改行コードの調整
        temp_content = content.replace('\r\n', '\n').replace('\r', '\n')
        final_content = temp_content.replace('\n', newline)

        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        
        with path.open('w', encoding=final_encoding, newline='') as f:
            f.write(final_content)

    @classmethod
    def batch_export(
        cls, 
        data: List[Dict[str, Any]], 
        template_proc: TemplateProcessor,
        dest_dir: str,
        filename_col: str,
        encoding: str,
        newline: str,
        overwrite: bool,
        combine_single: bool = False,
        separator: str = "---",
        output_filename: str = "infill_output.txt",
        progress_callback: Optional[Callable[[int, str, bool, str], None]] = None
    ) -> Tuple[int, int]:
        """
        複数行のデータを一括出力する。
        
        Returns:
            (成功数, 失敗数)
        """
        success_count = 0
        error_count = 0
        dest_path = Path(dest_dir)
        ext = template_proc.extension  # ひな形の拡張子を使用
        
        if combine_single:
            # 1つのファイルに結合する場合
            combined_content = []
            for i, row in enumerate(data):
                try:
                    content = template_proc.preview(row, highlight=False)
                    combined_content.append(content)
                    if progress_callback:
                        progress_callback(i, f"[準備] 行 {i+1} の処理中...", True, "")
                except Exception as e:
                    if progress_callback:
                        progress_callback(i, f"[エラー] 行 {i+1} の置換失敗: {e}", False, str(e))
            
            if not combined_content:
                return 0, 0

            # 指定されたセパレーターで結合（空の場合は改行のみ）
            if separator:
                sep_with_newlines = f"{newline}{separator}{newline}"
                final_text = sep_with_newlines.join(combined_content)
            else:
                final_text = newline.join(combined_content)
            
            # 指定されたファイル名を使用
            final_fn = output_filename.strip()
            if not final_fn:
                final_fn = f"combined_output{ext}"
            
            # 拡張子がない場合は補完
            if not Path(final_fn).suffix:
                final_fn += ext

            file_path = dest_path / final_fn
            if file_path.exists() and not overwrite:
                file_path = cls.get_unique_path(file_path)
            
            try:
                cls.export(final_text, str(file_path), encoding, newline)
                if progress_callback:
                    progress_callback(0, f"[成功] 結合ファイルを出力: {file_path.name}", True, "")
                return 1, 0
            except Exception as e:
                if progress_callback:
                    progress_callback(0, f"[失敗] 結合出力失敗: {e}", False, str(e))
                return 0, 1

        # 個別ファイル出力
        for i, row in enumerate(data):
            basename = ""
            if filename_col != "(連番出力)":
                basename = str(row.get(filename_col, "")).strip()
            
            if not basename:
                basename = f"output_{i+1:03d}"
            
            basename = cls.sanitize_filename(basename)
            if not basename:
                basename = f"output_{i+1:03d}"
            
            # 拡張子をひな形に合わせる
            file_path = dest_path / f"{basename}{ext}"
            
            log_msg = ""
            if file_path.exists() and not overwrite:
                old_name = file_path.name
                file_path = cls.get_unique_path(file_path)
                log_msg = f"[連番付与] {old_name} -> {file_path.name}"

            try:
                content = template_proc.preview(row, highlight=False)
                cls.export(content, str(file_path), encoding, newline)
                success_count += 1
                if progress_callback:
                    if log_msg: progress_callback(i, log_msg, True, "")
                    progress_callback(i, f"[成功] {file_path.name}", True, "")
            except Exception as e:
                error_count += 1
                if progress_callback:
                    progress_callback(i, f"[失敗] {file_path.name} (エラー: {e})", False, str(e))
        
        return success_count, error_count
