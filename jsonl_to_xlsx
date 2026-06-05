import json
import pandas as pd

# ============================================================
# JSONL 转 Excel 脚本
# 功能：读取 JSONL 文件，将其转换为 Excel (.xlsx) 格式
# 依赖：pandas, openpyxl
# 安装依赖：pip install pandas openpyxl
# ============================================================

# ---------------------------
# 【在此处修改文件名】
# input_file: 输入的 JSONL 文件路径（包含文件名和扩展名）
# output_file: 输出的 Excel 文件路径（包含文件名和扩展名）
# ---------------------------
input_file = "kb_mapping.jsonl"       # ← 修改为你的 JSONL 文件名
output_file = "kb_mapping.xlsx"      # ← 修改为你想要的 Excel 输出文件名


def jsonl_to_excel(jsonl_path, excel_path):
    """
    将 JSONL 文件转换为 Excel 文件

    参数:
        jsonl_path (str): JSONL 文件的路径
        excel_path (str): 输出 Excel 文件的路径
    """
    data = []

    # 逐行读取 JSONL 文件（每行是一个独立的 JSON 对象）
    try:
        with open(jsonl_path, "r", encoding="utf-8") as f:
            for line_number, line in enumerate(f, start=1):
                line = line.strip()
                # 跳过空行
                if not line:
                    continue
                try:
                    # 解析每一行的 JSON 数据
                    json_obj = json.loads(line)
                    data.append(json_obj)
                except json.JSONDecodeError as e:
                    print(f"⚠️  第 {line_number} 行解析失败，已跳过。错误信息: {e}")
    except FileNotFoundError:
        print(f"❌ 错误：找不到文件 '{jsonl_path}'，请检查文件路径是否正确。")
        return
    except Exception as e:
        print(f"❌ 读取文件时发生错误: {e}")
        return

    # 检查是否有有效数据
    if not data:
        print("❌ 错误：JSONL 文件中没有有效的数据。")
        return

    # 将数据转换为 DataFrame
    df = pd.DataFrame(data)

    # 写入 Excel 文件
    try:
        df.to_excel(excel_path, index=False, engine="openpyxl")
        print(f"✅ 转换成功！")
        print(f"   - 输入文件: {jsonl_path}")
        print(f"   - 输出文件: {excel_path}")
        print(f"   - 共转换 {len(df)} 行，{len(df.columns)} 列")
        print(f"   - 列名: {list(df.columns)}")
    except Exception as e:
        print(f"❌ 写入 Excel 文件时发生错误: {e}")


# 执行转换
if __name__ == "__main__":
    jsonl_to_excel(input_file, output_file)
