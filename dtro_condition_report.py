import argparse
import subprocess
import os


import mysql.connector

def fetch_unique_subprojects(project_id: int):
    connection = mysql.connector.connect(
        host="localhost",
        port=10645,
        user="deepinspector",
        password="xoaud17!@",
        database="db_deepinspector"
    )
    cursor = connection.cursor(dictionary=True)

    cursor.execute("""
        SELECT SUB_ID
        FROM SUB_PROJECT
        WHERE PROJECT_ID = %s;
    """, (project_id,))
    rows = cursor.fetchall()

    cursor.close()
    connection.close()

    return rows

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--root-dir", required=True)
    parser.add_argument("--project-id", required=True, type=int)
    parser.add_argument("--script-dir", required=True)
    args = parser.parse_args()

    # 1. 전체 보고서 실행
    subprocess.run([
        "python",
        os.path.join(args.script_dir, "dtro_total_statevl.py"),
        "--root-dir", args.root_dir,
        "--project-id", str(args.project_id)
    ])

    # 2. DB에서 가져왔다고 가정한 SUB_ID 목록
    sub_ids = fetch_unique_subprojects(args.project_id)
    print("sub_ids", sub_ids)

    # UP/DOWN 제거하고 중복 제거
    unique_prefixes = sorted({row["SUB_ID"].rsplit("_", 1)[0] for row in sub_ids})
    print("unique_prefixes", unique_prefixes)

    # 3. 각 sub-project 실행
    for prefix in unique_prefixes:
        subprocess.run([
            "python",
            os.path.join(args.script_dir, "dtro_dtl_statevl.py"),
            "--root-dir", args.root_dir,
            "--project-id", str(args.project_id),
            "--sub-project-id", prefix
        ])


if __name__ == "__main__":
    main()


