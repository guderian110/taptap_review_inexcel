import json
from pathlib import Path

this_dir = Path(__file__).resolve().parent
json_path = this_dir / "config.json"
response_path = this_dir / "response.json"

def load_config(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)  # 加载 JSON 数据
            return data
    except FileNotFoundError:
         print(f"文件未找到: {file_path}")
    except json.JSONDecodeError:
        print("文件内容不是有效的 JSON 格式")
    except Exception as e:
        print(f"发生错误: {e}")
        
def save_json_to_file(data, file_path):
    try:
        with open(file_path, 'w', encoding='utf-8') as file:
            json.dump(data, file, ensure_ascii=False, indent=4)  # 写入 JSON 数据
        print(f"数据已成功保存到 {file_path}")
    except Exception as e:
        print(f"发生错误: {e}")
        