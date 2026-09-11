"""
IRON LOG - ローカルWi-Fi配信サーバー (スマホ検証用)
PCとAndroidスマホが同じWi-Fiに接続されている状態で実行すると、
スマホのChromeから直接アプリを開くことができます。
"""
import http.server
import socket
import socketserver
import os
import sys

PORT = 8080

def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # Google DNSにダミー接続して自IPを取得
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

class Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        # キャッシュ無効ヘッダー（開発・検証用）
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

def run():
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
    ip = get_ip()
    url = f"http://{ip}:{PORT}/"
    local_url = f"http://localhost:{PORT}/"
    
    print("=" * 60)
    print("  IRON LOG - Android Chrome 配信サーバー")
    print("=" * 60)
    print(f"\n[1] PCブラウザで開く場合:\n    {local_url}")
    print(f"\n[2] AndroidスマホのChromeで開く場合:\n    {url}")
    print("\n※ スマホとPCが同じWi-Fi（ルーター）に接続されている必要があります。")
    print("※ スマホのChromeで上記URLを開き、右上のメニュー（縦の3点アイコン）から")
    print("   「ホーム画面に追加」をタップすると、全画面アプリとして登録されます。")
    print("\n終了するには Ctrl+C を押してください。\n" + "-" * 60)
    
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    socketserver.ThreadingTCPServer.allow_reuse_address = True
    with socketserver.ThreadingTCPServer(("", PORT), Handler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nサーバーを停止しました。")

if __name__ == '__main__':
    run()
