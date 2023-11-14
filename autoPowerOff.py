import wmi
import os

# WMIインスタンスを作成
w = wmi.WMI()

# シャットダウンを実行する関数
def shutdown():
    print("シャットダウンを実行します。")
    os.system("shutdown /s /f /t 0 /full")

# イベントを監視するWMIクエリ
watcher = w.Win32_NTLogEvent.watch_for("creation", EventCode='105')

print("イベント監視を開始します...")

try:
    while True:
        # イベントが生成されるのを待つ
        event = watcher()
        # InsertionStringsを使ってデータを取得する
        if event.InsertionStrings:
            ac_online = 'false' in event.InsertionStrings
            if ac_online:
                # AC電源がオフラインの場合、シャットダウンを実行
                shutdown()
                break

except KeyboardInterrupt:
    print("スクリプトが中断されました。")
