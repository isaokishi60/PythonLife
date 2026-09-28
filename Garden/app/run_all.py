import subprocess
import time
import os
import sys
import psutil


# このファイル(run_all.py)があるフォルダ
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def run_app(label: str, script_name: str, port: int):
    """指定した Streamlit アプリを別プロセスで起動する"""

    script_path = os.path.join(
        BASE_DIR,
        script_name
    )

    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        script_path,
        "--server.port",
        str(port),
    ]

    print(
        f"[INFO] {label} を起動します:"
    )
    print("       " + " ".join(cmd))

    try:
        p = subprocess.Popen(cmd)
        return p

    except Exception as e:
        print(
            f"[ERROR] {label} の起動に失敗しました: {e}"
        )
        return None


def kill_process_tree(
    proc: subprocess.Popen | None
):
    """
    proc とその子プロセスをまとめて終了する
    """

    if proc is None:
        return

    try:
        parent = psutil.Process(
            proc.pid
        )

    except psutil.NoSuchProcess:
        return

    # 子プロセスを先に停止
    children = parent.children(
        recursive=True
    )

    for child in children:

        try:
            child.terminate()

        except psutil.NoSuchProcess:
            pass

    # 親プロセスを停止
    try:
        parent.terminate()

    except psutil.NoSuchProcess:
        pass


def main():

    # ========================================================
    # Streamlitアプリ起動
    # ========================================================

    # メイン画面
    p_start = run_app(
        "start_page",
        "start_page.py",
        8501
    )

    # ガントチャート
    p_gantt = run_app(
        "ガント (sakutuke_gantt.py)",
        "sakutuke_gantt.py",
        8502
    )

    # layout_view.py は
    # start_page.py 内から呼び出すため、
    # 8503では起動しない

    time.sleep(3)

    print()
    print("=======================================")
    print(" 起動完了")
    print("=======================================")
    print()
    print("ブラウザで次のURLを開いてください：")
    print()
    print(
        "  http://localhost:8501  "
        "家庭菜園スタート画面"
    )
    print(
        "  http://localhost:8502  "
        "ガントビュー"
    )
    print()
    print(
        "レイアウトビューは"
        " start_page のボタンから表示します。"
    )
    print()
    print(
        "このウィンドウで Ctrl + C を押すと、"
        "すべて終了します。"
    )
    print()

    try:

        procs = [
            p
            for p in (
                p_start,
                p_gantt,
            )
            if p is not None
        ]

        for p in procs:
            p.wait()

    except KeyboardInterrupt:

        print()
        print(
            "[INFO] 終了処理中です…"
        )

        for p in (
            p_start,
            p_gantt,
        ):
            kill_process_tree(p)

        print(
            "[INFO] すべての "
            "Streamlit / Python プロセスを"
            "終了しました。"
        )

        sys.exit(0)


if __name__ == "__main__":
    main()

