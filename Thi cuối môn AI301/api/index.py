import os
import sys
import traceback

# Add root directory to sys.path so 'app' can be imported in Vercel Serverless environment
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

try:
    from app.main import app
except Exception as e:
    err_tb = traceback.format_exc()
    print("Vercel Serverless Startup Error:", err_tb)
    from flask import Flask
    app = Flask(__name__)

    @app.route("/", defaults={"path": ""})
    @app.route("/<path:path>")
    def vercel_error_handler(path):
        return f"""<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Khởi động ứng dụng Vercel</title>
    <style>
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #fff5f7; color: #881337; padding: 30px; }}
        .card {{ background: #ffffff; border: 1px solid #fecdd3; border-radius: 16px; padding: 24px; box-shadow: 0 10px 25px rgba(244,63,110,0.1); max-width: 900px; margin: 0 auto; }}
        h1 {{ color: #e11d48; font-size: 20px; margin-top: 0; }}
        pre {{ background: #fdf2f8; color: #9d174d; padding: 16px; border-radius: 12px; font-size: 12px; overflow-x: auto; border: 1px solid #fbcfe8; }}
        .btn {{ display: inline-block; background: #f43f6e; color: white; padding: 8px 18px; border-radius: 10px; text-decoration: none; font-weight: bold; font-size: 13px; margin-top: 15px; }}
    </style>
</head>
<body>
    <div class="card">
        <h1>⚠️ Thông báo Khởi động Máy chủ Vercel (Debug Log)</h1>
        <p>Ứng dụng đã được nạp thành công nhưng phát hiện lỗi nạp module trong môi trường Serverless của Vercel:</p>
        <pre>{err_tb}</pre>
        <p><b>Giải pháp:</b> Hãy kiểm tra file requirements.txt và các biến môi trường.</p>
        <a href="/" class="btn">Thử tải lại (Reload)</a>
    </div>
</body>
</html>""", 500

# Vercel looks for 'app' as the WSGI application callable
if __name__ == "__main__":
    app.run()
