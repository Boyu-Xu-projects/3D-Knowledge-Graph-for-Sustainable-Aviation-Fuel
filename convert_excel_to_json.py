# convert_excel_to_json.py
import pandas as pd
import json
import os

# ========= 配置 =========
input_path = r"C:\Users\xuboy\saf_llm_ner_project\KG-data\a-reaction-type-KG.xlsx"
output_dir = r"C:\Users\xuboy\saf_llm_ner_project\KG-data\kg_output"
sheet = 0   # 默认第一个工作表
# ======================

# 你定义的边规则（列名必须和 Excel 表头完全一致）
link_rules = [
    ("Year", "A reaction type"),
    ("A reaction type", "Feedstock"),
    ("Feedstock", "Operation mode"),
    ("Operation mode", "Catalyst"),
    ("Catalyst", "Product"),
    ("Product", "Product Selectivity"),
    ("Product", "Product yield"),
    ("Atmosphere", "Catalyst"),
    ("Reactant molar ratio", "Catalyst"),
    ("Flow rate", "Catalyst"),
    ("Reaction time", "Catalyst"),
    ("Reaction temperature", "Catalyst"),
    ("Reaction pressure", "Catalyst"),
    ("Solvent", "Catalyst"),
    ("Feedstock", "Conversion rate")
]

# 新增的category连接规则
category_link_rules = [
    ("Feedstock category", "Feedstock"),
    ("Catalyst category", "Catalyst"), 
    ("Product category", "Product")
]

os.makedirs(output_dir, exist_ok=True)

# 读取 Excel
df = pd.read_excel(input_path, sheet_name=sheet, engine="openpyxl")

nodes = []
links = []
node_map = {}  # 名称 -> id
next_id = 0

def get_node_id(name, ntype, category=None):
    """如果节点不存在就创建，记录它的类型和分类"""
    global next_id
    key = (name, ntype)
    if key not in node_map:
        node_data = {"id": next_id, "name": name, "type": ntype}
        if category:
            node_data["category"] = category
        node_map[key] = next_id
        nodes.append(node_data)
        next_id += 1
    return node_map[key]

# 首先处理category关系
for _, row in df.iterrows():
    for category_col, item_col in category_link_rules:
        if pd.notna(row.get(category_col)) and pd.notna(row.get(item_col)):
            category_name = str(row[category_col]).strip()
            item_name = str(row[item_col]).strip()
            if category_name and item_name:
                # 创建category节点
                category_id = get_node_id(category_name, f"{item_col.split()[0]} Category")  # 如 "Feedstock Category"
                # 创建item节点，并关联category
                item_id = get_node_id(item_name, item_col, category=category_name)
                
                # 添加category到item的连接
                links.append({
                    "source": category_id,
                    "target": item_id,
                    "relation": f"{category_col}->{item_col}"
                })

# 遍历 Excel 每一行，根据规则生成边
for _, row in df.iterrows():
    for source_col, target_col in link_rules:
        if pd.notna(row.get(source_col)) and pd.notna(row.get(target_col)):
            s_name = str(row[source_col]).strip()
            t_name = str(row[target_col]).strip()
            if s_name and t_name:
                s_id = get_node_id(s_name, source_col)
                t_id = get_node_id(t_name, target_col)
                links.append({
                    "source": s_id,
                    "target": t_id,
                    "relation": f"{source_col}->{target_col}"
                })

# 保存 JSON
with open(os.path.join(output_dir, "nodes.json"), "w", encoding="utf-8") as f:
    json.dump(nodes, f, ensure_ascii=False, indent=2)

with open(os.path.join(output_dir, "links.json"), "w", encoding="utf-8") as f:
    json.dump(links, f, ensure_ascii=False, indent=2)

with open(os.path.join(output_dir, "graph.json"), "w", encoding="utf-8") as f:
    json.dump({"nodes": nodes, "links": links}, f, ensure_ascii=False, indent=2)

print(f"导出完成: {len(nodes)} 个节点, {len(links)} 条边")
print("输出目录:", os.path.abspath(output_dir))

# 额外保存category信息用于前端分级显示
category_data = {
    "Feedstock": {
        "categories": list(set([n["category"] for n in nodes if n.get("type") == "Feedstock" and n.get("category")])),
        "items": {}
    },
    "Catalyst": {
        "categories": list(set([n["category"] for n in nodes if n.get("type") == "Catalyst" and n.get("category")])),
        "items": {}
    },
    "Product": {
        "categories": list(set([n["category"] for n in nodes if n.get("type") == "Product" and n.get("category")])),
        "items": {}
    }
}

# 组织每个category下的items
for node in nodes:
    if node.get("type") in ["Feedstock", "Catalyst", "Product"] and node.get("category"):
        category = node["category"]
        item_type = node["type"]
        if category not in category_data[item_type]["items"]:
            category_data[item_type]["items"][category] = []
        category_data[item_type]["items"][category].append(node["name"])

with open(os.path.join(output_dir, "categories.json"), "w", encoding="utf-8") as f:
    json.dump(category_data, f, ensure_ascii=False, indent=2)

print("Category数据导出完成")