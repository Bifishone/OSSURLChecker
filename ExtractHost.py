import os
import pandas as pd


def extract_host_from_xlsx(file_path):
    """从单个xlsx文件中提取Host列的内容"""
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)

        # 检查是否存在Host列
        if 'Host' not in df.columns:
            print(f"文件 {file_path} 中未找到Host列")
            return []

        # 提取Host列内容并去重
        host_list = df['Host'].dropna().unique().tolist()
        return host_list

    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {str(e)}")
        return []


def save_hosts_to_file(hosts, filename="url.txt"):
    """将host列表保存到文件中"""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            for host in hosts:
                f.write(f"{host}\n")
        print(f"已成功将 {len(hosts)} 个唯一Host保存到 {filename}")
    except Exception as e:
        print(f"保存文件时出错: {str(e)}")


def main():
    # 获取当前目录下所有的xlsx文件
    current_dir = os.getcwd()
    xlsx_files = [f for f in os.listdir(current_dir)
                  if f.endswith('.xlsx') and os.path.isfile(os.path.join(current_dir, f))]

    if not xlsx_files:
        print("当前目录下没有找到xlsx文件")
        return

    # 存储所有提取到的Host
    all_hosts = []

    # 处理每个xlsx文件
    for file in xlsx_files:
        file_path = os.path.join(current_dir, file)
        print(f"正在处理文件: {file}")

        hosts = extract_host_from_xlsx(file_path)
        all_hosts.extend(hosts)

        if hosts:
            print(f"从 {file} 中提取到 {len(hosts)} 个Host")
        print("---")

    # 去重并显示所有结果
    unique_hosts = list(set(all_hosts))
    print(f"所有文件中总共提取到 {len(unique_hosts)} 个唯一的Host")

    # 保存到文件
    save_hosts_to_file(unique_hosts)


if __name__ == "__main__":
    main()
