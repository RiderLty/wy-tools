import os
import json
import subprocess
import sys
from typing import Optional
import re
from typing import Dict, Optional
import pandas as pd
from typing import List, Dict, Any

def excel_to_dict_array(file_path: str) -> List[Dict[str, Any]]:
    """
    从 Excel 文件提取数据：第三行作为 key，第四行及之后每行作为一个字典。

    Args:
        file_path: Excel 文件路径（.xls 或 .xlsx）

    Returns:
        字典数组，每个字典代表一行数据（从第四行开始）

    Raises:
        FileNotFoundError: 文件不存在
        ValueError: 文件格式错误或行数不足
    """
    # 读取整个 Excel 文件，不将任何行作为列名
    df = pd.read_excel(file_path, header=None, dtype=str)  # 统一读取为字符串，避免类型问题

    if df.shape[0] < 4:
        raise ValueError("Excel 文件行数不足，至少需要 4 行（第三行作为 key，第四行开始作为数据）")

    # 第三行（索引 2）作为键
    keys = df.iloc[2].fillna('').astype(str).tolist()  # 将 NaN 转为空字符串

    # 第四行及之后（索引 3 到末尾）
    data_rows = df.iloc[3:].values

    result = []
    for row in data_rows:
        # 将 None/NaN 转为空字符串，然后转为列表
        row_values = [str(v) if pd.notna(v) else '' for v in row]
        # 若行长度不足 keys，用空字符串补齐；若超长则截断（按 keys 长度）
        if len(row_values) < len(keys):
            row_values.extend([''] * (len(keys) - len(row_values)))
        else:
            row_values = row_values[:len(keys)]
        # 构建字典
        row_dict = dict(zip(keys, row_values))
        result.append(row_dict)

    return result

def extract_first_match(text: str, patterns: Dict[str, str]) -> Dict[str, Optional[str]]:
    """
    对文本应用多个正则表达式，返回每个模式第一个匹配中的**第一个捕获组**内容。
    若模式中没有捕获组，则返回整个匹配。

    Args:
        text: 待搜索的纯文本字符串
        patterns: 字典，键为标识符，值为正则表达式模式（原始字符串）

    Returns:
        字典，键与 patterns 相同，值为捕获组内容（或整个匹配），未匹配则为 "null"
    """
    results = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            # 如果有捕获组，返回第一个捕获组的内容；否则返回整个匹配
            results[key] = match.group(1) if match.groups() else match.group(0)
        else:
            results[key] = "null"
    return results
def extract_text_from_doc(file_path: str) -> str:
    """
    从 .doc 或 .docx 文件中提取纯文本内容。

    Args:
        file_path: 文档路径

    Returns:
        提取的纯文本字符串

    Raises:
        FileNotFoundError: 文件不存在
        ValueError: 不支持的文件格式或缺少必要的工具/库
        RuntimeError: 提取过程中出现错误
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".docx":
        return _extract_docx(file_path)
    elif ext == ".doc":
        return _extract_doc(file_path)
    else:
        raise ValueError(f"不支持的文件类型: {ext}，仅支持 .doc 和 .docx")

def _extract_docx(file_path: str) -> str:
    """使用 python-docx 提取 .docx 文本"""
    try:
        from docx import Document
    except ImportError:
        raise RuntimeError("需要安装 python-docx 库: pip install python-docx")

    doc = Document(file_path)
    paragraphs = [para.text for para in doc.paragraphs]
    return "\n".join(paragraphs)

def _extract_doc(file_path: str) -> str:
    """尝试多种方式提取 .doc 文本"""
    # 尝试 1: 使用 catdoc (Linux/macOS/Windows 可用)
    if _command_exists("catdoc"):
        try:
            result = subprocess.run(
                ["catdoc", file_path],
                capture_output=True,
                text=True,
                check=True,
                encoding="utf-8"
            )
            return result.stdout
        except (subprocess.SubprocessError, UnicodeDecodeError):
            pass  # 失败则继续尝试其他方法

    # 尝试 2: 使用 antiword (仅支持 Word 97-2003)
    if _command_exists("antiword"):
        try:
            result = subprocess.run(
                ["antiword", file_path],
                capture_output=True,
                text=True,
                check=True,
                encoding="utf-8"
            )
            return result.stdout
        except subprocess.SubprocessError:
            pass

    # 尝试 3: Windows 下使用 win32com (需安装 pywin32 和 Word)
    if sys.platform == "win32":
        try:
            import win32com.client
        except ImportError:
            pass
        else:
            try:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(file_path)
                text = doc.Content.Text
                doc.Close()
                word.Quit()
                return text
            except Exception as e:
                raise RuntimeError(f"win32com 提取失败: {e}")

    # 所有方法都失败
    raise RuntimeError(
        "无法提取 .doc 文件内容。请安装以下任一工具：\n"
        "  - catdoc (推荐) : https://github.com/techtipsy/catdoc\n"
        "  - antiword      : http://www.winfield.demon.nl/\n"
        "  - Windows 用户可安装 pywin32 并确保已安装 Microsoft Word"
    )

def _command_exists(cmd: str) -> bool:
    """检查命令是否在 PATH 中可用"""
    return any(
        os.access(os.path.join(path, cmd), os.X_OK)
        for path in os.environ.get("PATH", "").split(os.pathsep)
        if os.path.exists(os.path.join(path, cmd))
    )

if __name__ == "__main__":
    try:
        text = extract_text_from_doc(r"1 来安县财政工程造价审核工作领导小组2026年第1次会议议题.docx")
        # print(text)
        word_proj_data = {}
        for single_proj in text.split("项目基本情况"):
            single_proj = single_proj.strip()
            if not single_proj:
                continue
            patterns = {
                "项目名称": r"项目名称：(\S+)",
                "建设单位": r"建设单位：(\S+)",
                "监理单位": r"监理单位：(\S+)",
                "施工单位": r"施工单位：(\S+)",
                "初审单位": r"初审单位：(\S+)",
                "复审单位": r"复审单位：(\S+)",
                "建设单位报审价":r"(?:建设单位|施工单位)报审价：(\d+(?:\.\d+)?)元",
                "初审审核价": r"初审审核价：(\d+(?:\.\d+)?)元",
                "初审核减金额": r"初审核减金额：(\d+(?:\.\d+)?)元",
                "初审核减率": r"初审核减率：(\d+(?:\.\d+)?)%",
                "复审审核价": r"复审审核价：(\d+(?:\.\d+)?)元",
                "复审核减金额": r"复审核减金额：(\d+(?:\.\d+)?)元",
                "复审核减率": r"复审核减率：(\d+(?:\.\d+)?)%",
                "总核减率": r"总核减率：(\d+(?:\.\d+)?)%",
                "一审误差率": r"工程造价一审误差率为(\d+(?:\.\d+)?)%",
            }

            result = extract_first_match(single_proj.replace(" ", "").replace("\t", ""), patterns)
            if not result:
                continue
            if not result["项目名称"]:
                continue
            if result["项目名称"] == "null":
                continue
            if "null" in json.dumps(result, ensure_ascii=False, indent=4):
                print("====================================")
                print(json.dumps(result, ensure_ascii=False, indent=4))
                print(single_proj.replace(" ", "").replace("\t", "").replace(":","："))
            else:
                pass
            word_proj_data[result["项目名称"].strip()] = result
        data = excel_to_dict_array(r"来安县财政工程造价审核工作领导小组2026年第1次会议明细表.xls")
        print("\n\n\n")
        
        for sg in data:
            if sg["项目名称"].strip() in word_proj_data:
                # print(word_proj_data[sg["项目名称"].strip()])
                pass
            else:
                print("未找到  "  + sg["项目名称"].strip() )
    except Exception as e:
        print(f"提取文件内容时出错: {e}")
        pass
    print("\n\n\n")
    for sg in data:
        if sg["项目名称"].strip() in word_proj_data:
            try:
                # print(json.dumps(word_proj_data[sg["项目名称"].strip()], ensure_ascii=False, indent=4))
                # print("\n")
                # print(json.dumps(sg, ensure_ascii=False, indent=4))
                # print("\n\n\n")
                w_data = word_proj_data[sg["项目名称"].strip()]
                e_data = sg
                
                if w_data["建设单位"].strip() != e_data["建设单位\n（业主）"].strip():
                    print(f"{w_data['项目名称']} 建设单位不一致: [{w_data['建设单位']}] != [{e_data['建设单位\n（业主）']}]")
                
                if w_data["监理单位"].strip() != e_data["监理单位"].strip():
                    print(f"{w_data['项目名称']} 监理单位不一致: [{w_data['监理单位']}] != [{e_data['监理单位']}]")

                if w_data["施工单位"].strip() != e_data["承包（施工）\n单位"].strip():
                    print(f"{w_data['项目名称']} 施工单位不一致: [{w_data['施工单位']}] != [{e_data['承包（施工）\n单位']}]")
                
                if w_data["初审单位"].strip() != e_data["结算审计单位（一审）"].strip():
                    print(f"{w_data['项目名称']} 初审单位不一致: [{w_data['初审单位']}] != [{e_data['结算审计单位（一审）']}]")
                
                if w_data["复审单位"].strip() != e_data["审查复核单位（复审）"].strip():
                    print(f"{w_data['项目名称']} 复审单位不一致: [{w_data['复审单位']}] != [{e_data['审查复核单位（复审）']}]")


                if abs(float(w_data["建设单位报审价"].strip()) - float(e_data["报审价\n（万元）"].strip()) * 10000) > 0.01:
                    print(f"{w_data['项目名称']} 建设单位报审价不一致: [{w_data['建设单位报审价']}] != [{float(e_data['报审价\n（万元）'].strip()) * 10000}]")

                if abs(float(w_data["初审审核价"].strip()) - float(e_data["一审审核价\n（万元）"].strip()) * 10000) > 0.01:
                    print(f"{w_data['项目名称']} 初审审核价不一致: [{w_data['初审审核价']}] != [{float(e_data['一审审核价\n（万元）'].strip()) * 10000}]")

                if abs(float(w_data["初审核减金额"].strip()) - float(e_data["一审核减造价\n（万元）"].strip()) * 10000) > 0.01:
                    print(f"{w_data['项目名称']} 初审核减金额不一致: [{w_data['初审核减金额']}] != [{float(e_data['一审核减造价\n（万元）'].strip()) * 10000}]")
                
                if abs(float(w_data["初审核减率"].strip()) - float(e_data["初审核减率"].strip()) * 100) > 0.01:
                    print(f"{w_data['项目名称']} 初审核减率不一致: [{w_data['初审核减率']}] != [{float(e_data['初审核减率'].strip()) * 100}]")
                
                if abs(float(w_data["复审审核价"].strip()) - float(e_data["复审价\n（万元）"].strip()) * 10000) > 0.01:
                    print(f"{w_data['项目名称']} 复审审核价不一致: [{w_data['复审审核价']}] != [{float(e_data['复审价\n（万元）'].strip()) * 10000}]")

                if abs(float(w_data["复审核减金额"].strip()) - float(e_data["核减造价\n（万元）"].strip()) * 10000) > 0.01:
                    print(f"{w_data['项目名称']} 复审核减金额不一致: [{w_data['复审核减金额']}] != [{float(e_data['核减造价\n（万元）'].strip()) * 10000}]")

                if abs(float(w_data["复审核减率"].strip()) - float(e_data["复审核减率"].strip()) * 100) > 0.01:
                    print(f"{w_data['项目名称']} 复审核减率不一致: [{w_data['复审核减率']}] != [{float(e_data['复审核减率'].strip()) * 100}]")

                if abs(float(w_data["总核减率"].strip()) - float(e_data["总核减率"].strip()) * 100) > 0.01:
                    print(f"{w_data['项目名称']} 总核减率不一致: [{w_data['总核减率']}] != [{float(e_data['总核减率'].strip()) * 100}]")
            
                if abs(float(w_data["一审误差率"].strip()) - float(e_data["一审误差率"].strip()) * 100) > 0.01:
                    print(f"{w_data['项目名称']} 一审误差率不一致: [{w_data['一审误差率']}] != [{float(e_data['一审误差率'].strip()) * 100}]")

            
            except Exception as e:
                print(f"{w_data['项目名称']} 异常: {e}")
                pass
    input("any key to continue")