import subprocess
import time
import os
import sys
import psutil  # 子プロセスごと安全に殺すために使用


# このファイル(run_all.py)があるフォルダ
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def run_app(label: str, script_name: str, port: int):
    """指定した Streamlit アプリを別プロセスで起動する"""
    script_path = os.path.join(BASE_DIR, script_name)
    cmd = [
        "streamlit",
        "run",
        script_path,
        "--server.port", str(port),
    ]
    print(f"[INFO] {label} を起動します: {' '.join(cmd)}")
    try:
        p = subprocess.Popen(cmd)
        return p
    except Exception as e:
        print(f"[ERROR] {label} の起動に失敗しました: {e}")
        return None


def kill_process_tree(proc: subprocess.Popen | None):
    """proc と、その子プロセス（python/streamlit）をまとめて終了させる"""
    if proc is None:
        return
    try:
        parent = psutil.Process(proc.pid)
    except psutil.NoSuchProcess:
        return

    # 子プロセスを先に止める
    children = parent.children(recursive=True)
    for child in children:
        try:
            child.terminate()
        except psutil.NoSuchProcess:
            pass

    # 親も止める
    try:
        parent.terminate()
    except psutil.NoSuchProcess:
        pass


def main():
    # ① 各アプリを起動（★ポート割り当てを変更）
    p_start = run_app("start_page (カレンダー)", "start_page.py", 8501)
    p_gantt = run_app("ガント (sakutuke_gantt.py)", "sakutuke_gantt.py", 8502)
    p_layout = run_app("レイアウト (layout_view.py)", "layout_view.py", 8503)

    time.sleep(3)

    print("\n=== 起動完了 ===")
    print("ブラウザで次のURLを開いてください：")
    print("  http://localhost:8501  start_page")
    print("  http://localhost:8502  ガントビュー")
    print("  http://localhost:8503  レイアウトビュー")
    print("\nこのウィンドウで Ctrl + C を押すと、3つのアプリをすべて終了します。\n")

    try:
        procs = [p for p in (p_start, p_gantt, p_layout) if p is not None]
        for p in procs:
            p.wait()
    except KeyboardInterrupt:
        print("\n[INFO] 終了処理中です…")
        for p in (p_start, p_gantt, p_layout):
            kill_process_tree(p)
        print("[INFO] すべての Streamlit / Python プロセスを終了しました。")
        sys.exit(0)



if __name__ == "__main__":
    main()


