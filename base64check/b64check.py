#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import json
import base64
import re
import xml.etree.ElementTree as ET

# ---------- 日志文件路径 ----------
LOG_FILE = "console_log.txt"
DIFF_REPORT_FILE = "diff_report.txt"
XML_ERROR_FILE = "xml_decode_errors.txt"  # 专门记录 XML 解码失败

# 打开日志和错误记录文件
log_fp = open(LOG_FILE, "w", encoding="utf-8")
xml_error_fp = open(XML_ERROR_FILE, "w", encoding="utf-8")


# --------- Logger 类：同时写到控制台和 log_fp ----------
class Logger:
    def __init__(self, *files):
        self.files = files

    def write(self, message):
        for f in self.files:
            try:
                f.write(message)
            except ValueError:
                pass

    def flush(self):
        for f in self.files:
            try:
                f.flush()
            except:
                pass


sys.stdout = Logger(sys.stdout, log_fp)


# ---------- 工具函数 ----------

def minify_json(content: str) -> str:
    """加载 JSON 并 dump 为紧凑格式。"""
    data = json.loads(content)
    return json.dumps(data, ensure_ascii=False, separators=(",", ":"))


def decode_xml_value(raw: str) -> str:
    """Base64 解码并还原 Unicode 转义。"""
    b = base64.b64decode(raw)
    return b.decode("utf-8").encode("utf-8").decode("unicode-escape")


def find_string_in_xml(xml_path: str, key: str) -> str:
    """在 strings.xml 中找 name=key 的 <string> 节点文本。"""
    tree = ET.parse(xml_path)
    for elem in tree.getroot().findall("string"):
        if elem.get("name") == key:
            return elem.text or ""
    raise KeyError(f"未在 {xml_path} 找到 key='{key}'")


def clean_value(value: str) -> str:
    """
    通用清洗：
    1. "\\n" → 换行, "\\t" → 制表符；
    2. 统一换行符为 "\n"；
    3. 删除剩余反斜杠；
    4. 去除首尾双引号及空白；
    5. 分行去首尾空白，丢空行，用单空格拼接。
    """
    value = value.replace("\\n", "\n").replace("\\t", "\t")
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    value = value.replace("\\", "")
    value = value.strip('"').strip()
    lines = [line.strip() for line in value.split("\n") if line.strip()]
    return " ".join(lines)


def normalize_ws_preserve_tab(s: str) -> str:
    """合并空格/换行为单空格，保留制表符\t。"""
    tmp = s.replace("\n", " ")
    tmp = re.sub(r" {2,}", " ", tmp)
    return tmp.strip()


# ---------- 主流程 ----------

def main():
    print("脚本开始运行...\n")
    # 写 diff report 头部
    with open(DIFF_REPORT_FILE, "w", encoding="utf-8") as diff_fp:
        diff_fp.write("JSON vs XML 解码不一致详尽报告\n")
        diff_fp.write("=" * 60 + "\n\n")

        # 遍历当前目录
        for dirpath, _, files in os.walk("."):
            # 跳过资源文件夹
            if "res" in dirpath.split(os.sep):
                continue
            # 只处理语言代码文件夹
            parts = dirpath.strip(os.sep).split(os.sep)
            if not parts or not re.fullmatch(r"[a-z]{2,3}", parts[-1]):
                continue
            lang = parts[-1]

            for fname in files:
                if not fname.endswith(".json"):
                    continue

                json_path = os.path.join(dirpath, fname)
                # 1. 压缩 JSON
                try:
                    raw_json = open(json_path, encoding="utf-8").read()
                    jmin = minify_json(raw_json)
                    print(f"已压缩 JSON: {json_path}")
                except Exception as e:
                    print(f"⚠️ JSON 解析失败: {json_path}  错误: {e}")
                    continue

                # 2. 定位 strings.xml
                xml_dir = os.path.join("res", "values" if lang == "en" else f"values-{lang}")
                xml_path = os.path.join(xml_dir, "strings.xml")
                if not os.path.isfile(xml_path):
                    print(f"⚠️ 未找到 XML: {xml_path}，跳过")
                    continue

                key = fname[:-5]
                # 3. 取 base64 文本
                try:
                    raw_xml = find_string_in_xml(xml_path, key)
                    print(f"找到 XML key: {key} 在 {xml_path}")
                except KeyError as e:
                    print(f"⚠️ {e}")
                    continue

                # 4. 解码 XML
                try:
                    xraw = decode_xml_value(raw_xml)
                except Exception as e:
                    msg = f"XML 解码失败: {xml_path}#{key}  错误: {e}"
                    print(f"❌ {msg}")
                    # 记录到 xml_decode_errors.txt
                    xml_error_fp.write(msg + "\n")
                    continue

                # 5. 清洗 & 空白归一
                jclean = normalize_ws_preserve_tab(clean_value(jmin))
                xclean = normalize_ws_preserve_tab(clean_value(xraw))

                # 6. 比对并写入 diff report
                if jclean != xclean:
                    idx = next(
                        (i for i, (a, b) in enumerate(zip(jclean, xclean)) if a != b),
                        min(len(jclean), len(xclean))
                    )
                    ctx = 30
                    s = max(idx - ctx, 0)
                    ej = min(idx + ctx + 1, len(jclean))
                    ex = min(idx + ctx + 1, len(xclean))

                    diff_fp.write(f"文件: {json_path}\n")
                    diff_fp.write(f"对应: {xml_path}#{key}\n")
                    diff_fp.write(f"位置 (0-base): {idx}\n")
                    diff_fp.write(
                        f"  JSON[{idx}]: '{jclean[idx] if idx < len(jclean) else 'EOF'}'  vs  "
                        f"XML[{idx}]: '{xclean[idx] if idx < len(xclean) else 'EOF'}'\n\n"
                    )
                    diff_fp.write("【清洗后 JSON 片段】\n")
                    diff_fp.write(f"…{jclean[s:ej]}…\n\n")
                    diff_fp.write("【清洗后 XML 片段】\n")
                    diff_fp.write(f"…{xclean[s:ex]}…\n\n")
                    diff_fp.write("-" * 60 + "\n\n")
                    print(f"❌ 不一致: {json_path} ↔ {xml_path}#{key}")
                else:
                    print(f"✅ 一致: {json_path} ↔ {xml_path}#{key}")

    # 脚本结束提示
    print(f"\n所有处理完成。")
    print(f"差异报告：{os.path.abspath(DIFF_REPORT_FILE)}")
    print(f"XML 解码失败记录：{os.path.abspath(XML_ERROR_FILE)}")
    print(f"控制台日志：{os.path.abspath(LOG_FILE)}")

    # 不显式 close，Python 退出时自动关闭
    # log_fp.close()
    # xml_error_fp.close()


if __name__ == "__main__":
    main()
