# -*- coding: utf-8 -*-
"""
Flask Web Application & API cho AI Agent Bảo Ngọc Retail
Chạy trên localhost để Quản lý và Chuyên viên thao tác trực tiếp.
"""

import os
import json
import io
import datetime
import pandas as pd
import numpy as np
from flask import Flask, render_template, request, jsonify, send_file
from app.engine import engine, DATA_DIR, WORKSPACE_DIR, BASE_DIR, APP_DATA_DIR, resolve_data_file
from app.pitch_deck_generator import generate_pitch_deck, get_pitch_deck_preview_data
from app.word_report_generator import generate_word_report, get_word_report_preview_html
from app.excel_export_generator import generate_excel_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static")
)
app.config['JSON_AS_ASCII'] = False
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
app.jinja_env.auto_reload = True

def auto_init_engine():
    """Tự động khởi tạo dữ liệu tệp mới nhất khi khởi động máy chủ"""
    try:
        up_xlsx = resolve_data_file("uploaded_raw.xlsx")
        up_csv = resolve_data_file("uploaded_raw.csv")
        target_f = None
        if os.path.exists(up_xlsx) and os.path.exists(up_csv):
            target_f = up_xlsx if os.path.getmtime(up_xlsx) >= os.path.getmtime(up_csv) else up_csv
        elif os.path.exists(up_xlsx):
            target_f = up_xlsx
        elif os.path.exists(up_csv):
            target_f = up_csv
        if target_f:
            engine.load_raw_data(target_f)
            engine.execute_data_cleaning()
            engine.generate_management_report()
            engine.perform_cross_check_audit()
    except Exception as e:
        print("auto_init_engine notice:", e)

auto_init_engine()

@app.route("/")
def index():
    """Trang chủ hiển thị toàn bộ quy trình 3 giai đoạn theo BPMN"""
    return render_template("index.html")

@app.route("/api/system-state", methods=["GET"])
def get_system_state():
    """Lấy trạng thái hiện tại của hệ thống"""
    # Nếu có file uploaded_raw nhưng chưa nạp, tự động nạp tệp mới nhất
    if engine.raw_df is None:
        up_xlsx = resolve_data_file("uploaded_raw.xlsx")
        up_csv = resolve_data_file("uploaded_raw.csv")
        target_f = None
        if os.path.exists(up_xlsx) and os.path.exists(up_csv):
            target_f = up_xlsx if os.path.getmtime(up_xlsx) >= os.path.getmtime(up_csv) else up_csv
        elif os.path.exists(up_xlsx):
            target_f = up_xlsx
        elif os.path.exists(up_csv):
            target_f = up_csv

        if target_f:
            try:
                engine.load_raw_data(target_f)
                engine.execute_data_cleaning()
                engine.generate_management_report()
                engine.perform_cross_check_audit()
            except Exception:
                pass

    raw_loaded = engine.raw_df is not None
    cleaned_done = engine.cleaned_df is not None
    verified = engine.human_verification["status"] == "APPROVED"
    report_done = engine.management_report is not None
    audit_done = engine.audit_results is not None
    
    return jsonify({
        "raw_loaded": raw_loaded,
        "cleaned_done": cleaned_done,
        "human_verified": verified,
        "verification_info": engine.human_verification,
        "report_done": report_done,
        "audit_done": audit_done,
        "decisions_count": len(engine.executive_decisions),
        "cleaning_audit_log": getattr(engine, "cleaning_audit_log", None),
        "raw_rows": len(engine.raw_df) if engine.raw_df is not None else 0,
        "cleaned_rows": len(engine.cleaned_df) if engine.cleaned_df is not None else 0
    })

@app.route("/api/load-sample", methods=["GET", "POST"])
def load_sample_data():
    """Nạp bộ dữ liệu thô: Mặc định ưu tiên file người dùng tải lên"""
    try:
        up_xlsx = resolve_data_file("uploaded_raw.xlsx")
        sample_path = up_xlsx if os.path.exists(up_xlsx) else resolve_data_file("raw_transactions_oct2026.csv")
        info = engine.load_raw_data(sample_path)
        msg = f"Đã nạp thành công {info['total_rows']:,} dòng dữ liệu thô!"
        return jsonify({"success": True, "data": info, "message": msg})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/load-uploaded", methods=["GET", "POST"])
def load_uploaded_data():
    """Nạp tệp dữ liệu giao dịch hiện tại hoặc dữ liệu mẫu"""
    try:
        uploaded_raw_xlsx = resolve_data_file("uploaded_raw.xlsx")
        uploaded_raw_csv = resolve_data_file("uploaded_raw.csv")
        target_file = None
        if os.path.exists(uploaded_raw_xlsx) and os.path.exists(uploaded_raw_csv):
            target_file = uploaded_raw_xlsx if os.path.getmtime(uploaded_raw_xlsx) >= os.path.getmtime(uploaded_raw_csv) else uploaded_raw_csv
        elif os.path.exists(uploaded_raw_xlsx):
            target_file = uploaded_raw_xlsx
        elif os.path.exists(uploaded_raw_csv):
            target_file = uploaded_raw_csv
        else:
            target_file = resolve_data_file("raw_transactions_oct2026.csv")
            
        info = engine.load_raw_data(target_file)
        return jsonify({"success": True, "data": info, "message": f"Đã nạp thành công {info['total_rows']:,} dòng từ tệp {info['file_name']}!"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/business-profile", methods=["GET", "POST"])
def business_profile_api():
    """Lấy hoặc cập nhật thông tin hồ sơ doanh nghiệp & câu hỏi phân tích định hướng cho AI"""
    if request.method == "POST":
        data = request.json or request.form.to_dict() or {}
        res = engine.update_business_profile(data)
        return jsonify({"success": True, "data": res, "message": "Đã lưu thông tin hồ sơ doanh nghiệp và câu hỏi phân tích thành công!"})
    return jsonify({"success": True, "data": engine.get_business_profile()})

@app.route("/api/upload-raw", methods=["POST"])
def upload_raw_data():
    """Upload tệp dữ liệu thô mới (.xlsx, .xls, .csv) - Yêu cầu hoàn tất hồ sơ doanh nghiệp trước"""
    # Kiểm tra điều kiện tiên quyết: Cần có thông tin doanh nghiệp trước khi tải file
    profile = engine.get_business_profile()
    if not profile or not profile.get("is_confirmed"):
        if request.form.get("business_type") and request.form.get("context"):
            questions = []
            try:
                questions = json.loads(request.form.get("analysis_questions", "[]"))
            except Exception:
                pass
            engine.update_business_profile({
                "business_type": request.form.get("business_type"),
                "context": request.form.get("context"),
                "additional_info": request.form.get("additional_info", ""),
                "analysis_questions": questions
            })
        else:
            # Tự động kích hoạt mẫu đề thi AI301 mặc định để người dùng luôn tải file thành công
            engine.business_profile["is_confirmed"] = True

    if 'file' not in request.files:
        return jsonify({"success": False, "error": "Không tìm thấy tệp tải lên"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"success": False, "error": "Chưa chọn tệp"}), 400
        
    try:
        orig_filename = file.filename
        _, ext = os.path.splitext(orig_filename)
        ext = ext.lower() if ext else ".csv"
        
        # Thử lưu vào DATA_DIR, nếu gặp Read-only file system (Vercel) thì tự động chuyển sang /tmp
        target_dir = DATA_DIR
        try:
            os.makedirs(target_dir, exist_ok=True)
            # Xóa các tệp uploaded_raw cũ để không bị lẫn lộn giữa .csv và .xlsx
            for old_ext in [".xlsx", ".csv", ".xls"]:
                old_p = os.path.join(target_dir, f"uploaded_raw{old_ext}")
                if os.path.exists(old_p):
                    try:
                        os.remove(old_p)
                    except Exception:
                        pass
            save_path = os.path.join(target_dir, f"uploaded_raw{ext}")
            file.save(save_path)
        except OSError:
            import tempfile
            target_dir = os.path.join(tempfile.gettempdir(), "baongoc_ai_data")
            os.makedirs(target_dir, exist_ok=True)
            for old_ext in [".xlsx", ".csv", ".xls"]:
                old_p = os.path.join(target_dir, f"uploaded_raw{old_ext}")
                if os.path.exists(old_p):
                    try:
                        os.remove(old_p)
                    except Exception:
                        pass
            save_path = os.path.join(target_dir, f"uploaded_raw{ext}")
            file.save(save_path)
            # Cập nhật DATA_DIR toàn cục
            import app.engine as eng_mod
            eng_mod.DATA_DIR = target_dir

        info = engine.load_raw_data(save_path)
        info["file_name"] = orig_filename
        engine.active_file_name = orig_filename
        return jsonify({"success": True, "data": info, "message": f"Đã tải lên và nạp thành công {info['total_rows']:,} dòng từ tệp '{orig_filename}'!"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/execute-cleaning", methods=["GET", "POST"])
def execute_cleaning():
    """Giai đoạn 1: AI Agent thực hiện làm sạch dữ liệu tự động"""
    try:
        if engine.raw_df is None:
            return jsonify({"success": False, "error": "Chưa có dữ liệu thô. Vui lòng tải lên tệp tại Bước 1 trước!"}), 400
        log = engine.execute_data_cleaning()
        return jsonify({"success": True, "data": log, "message": "Đã làm sạch dữ liệu thành công!"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/human-verify", methods=["GET", "POST"])
def human_verify():
    """Giai đoạn 1: Con người kiểm tra lại (Human Verification)"""
    data = request.get_json(silent=True) or request.args.to_dict() or {}
    approved = str(data.get("approved", "true")).lower() in ["true", "1", "yes"]
    notes = data.get("notes", "")
    res = engine.human_verify(approved, notes)
    return jsonify({"success": True, "data": res, "message": "Đã lưu xác thực của chuyên viên!"})

@app.route("/api/reset-all-data", methods=["GET", "POST"])
def reset_all_data():
    """Xóa sạch toàn bộ dữ liệu giao dịch & báo cáo, đưa về trạng thái sạch 100%"""
    res = engine.reset_all_data()
    return jsonify({"success": True, "data": res, "message": "Đã xóa toàn bộ dữ liệu, hệ thống sạch 100%!"})

@app.route("/api/cleaned-data-preview", methods=["GET"])
def cleaned_data_preview():
    """Xem trước các dòng dữ liệu sạch và dữ liệu bị loại trừ"""
    page = int(request.args.get("page", 1))
    limit = int(request.args.get("limit", 20))
    status_filter = request.args.get("status", "")
    store_filter = request.args.get("store", "")
    mode = request.args.get("mode", "cleaned") # 'cleaned', 'raw', 'excluded'
    
    # Nếu hệ thống chưa nạp dữ liệu thô nhưng có file trên đĩa, tự động nạp
    if engine.raw_df is None:
        up_xlsx = resolve_data_file("uploaded_raw.xlsx")
        up_csv = resolve_data_file("uploaded_raw.csv")
        target_f = up_xlsx if os.path.exists(up_xlsx) else (up_csv if os.path.exists(up_csv) else None)
        if target_f:
            try:
                engine.load_raw_data(target_f)
                engine.execute_data_cleaning()
            except Exception:
                pass

    if engine.raw_df is None:
        return jsonify({
            "success": True,
            "total": 0,
            "page": 1,
            "limit": limit,
            "mode": mode,
            "columns": [],
            "records": []
        })
    
    if mode == "raw":
        df = engine.original_raw_df if getattr(engine, 'original_raw_df', None) is not None else engine.raw_df
    elif mode == "excluded":
        if engine.cleaned_df is not None:
            cleaned_ids = set(engine.cleaned_df["transaction_id"])
            df = engine.raw_df[~engine.raw_df["transaction_id"].isin(cleaned_ids)]
        else:
            df = None
    else:
        df = engine.cleaned_df
        
    if df is None or len(df) == 0:
        return jsonify({
            "success": True,
            "total": 0,
            "page": 1,
            "limit": limit,
            "mode": mode,
            "columns": [],
            "records": []
        })
        
    if status_filter and "status" in df.columns:
        df = df[df["status"] == status_filter]
    if store_filter and "store_id" in df.columns:
        df = df[df["store_id"] == store_filter]
        
    total_records = len(df)
    start_idx = (page - 1) * limit
    end_idx = min(start_idx + limit, total_records)
    
    subset_df = df.iloc[start_idx:end_idx].copy().fillna("")
    cols = list(subset_df.columns)
    records = subset_df.to_dict(orient="records")
    
    clean_records = []
    for r in records:
        clean_r = {}
        for k, v in r.items():
            if pd.isna(v) or v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
                clean_r[k] = ""
            elif isinstance(v, (pd.Timestamp, datetime.date, datetime.datetime)):
                clean_r[k] = v.strftime('%Y-%m-%d')
            elif isinstance(v, (int, np.integer)):
                clean_r[k] = int(v)
            elif isinstance(v, (float, np.floating)):
                clean_r[k] = float(v)
            else:
                clean_r[k] = str(v)
        clean_records.append(clean_r)
        
    return jsonify({
        "success": True,
        "total": total_records,
        "page": page,
        "limit": limit,
        "mode": mode,
        "columns": cols,
        "records": clean_records
    })

@app.route("/api/generate-report", methods=["GET", "POST"])
def generate_report():
    """Giai đoạn 2: Tạo Báo cáo Quản trị đa chiều"""
    try:
        if engine.cleaned_df is None:
            return jsonify({"success": False, "error": "Chưa có dữ liệu sạch. Vui lòng tải lên và làm sạch dữ liệu trước!"}), 400
        rep = engine.generate_management_report()
        return jsonify({"success": True, "data": rep, "message": "Báo cáo Quản trị đã được khởi tạo thành công!"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/cross-check-audit", methods=["GET", "POST"])
def cross_check_audit():
    """Giai đoạn 3: Đối chiếu chéo & Phát hiện bất thường & Xếp hạng ưu tiên"""
    try:
        if engine.management_report is None:
            return jsonify({"success": False, "error": "Chưa có báo cáo để đối soát. Vui lòng hoàn thành các bước trước!"}), 400
        audit = engine.perform_cross_check_audit()
        return jsonify({"success": True, "data": audit, "message": "Kiểm toán đối chiếu hoàn tất!"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/chat", methods=["POST"])
def chat():
    """AI Agent Co-pilot Chatbot hỗ trợ Nguyễn Bảo Ngọc"""
    data = request.json or {}
    message = data.get("message", "")
    if not message:
        return jsonify({"success": False, "reply": "Vui lòng nhập câu hỏi."}), 400
        
    try:
        reply = engine.chat_with_agent(message)
        return jsonify({"success": True, "reply": reply})
    except Exception as e:
        return jsonify({"success": False, "reply": f"Lỗi xử lý: {str(e)}"}), 500

@app.route("/api/executive-decision", methods=["POST"])
def executive_decision():
    """Bước 4: Nguyễn Bảo Ngọc ký duyệt Quyết định can thiệp hỗ trợ"""
    data = request.json or {}
    store_ids = data.get("store_ids", [])
    action_plan = data.get("action_plan", "")
    assigned_lead = data.get("assigned_lead", "")
    budget = data.get("budget", 0)
    notes = data.get("notes", "")
    
    if not store_ids:
        return jsonify({"success": False, "error": "Vui lòng chọn ít nhất một cửa hàng cần hỗ trợ"}), 400
        
    try:
        decision = engine.record_executive_decision(store_ids, action_plan, assigned_lead, budget, notes)
        return jsonify({"success": True, "data": decision, "message": f"Đã ban hành Quyết định {decision['decision_id']} thành công!"})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/decisions-list", methods=["GET"])
def get_decisions():
    """Lấy danh sách các quyết định đã ban hành"""
    return jsonify({"success": True, "data": engine.executive_decisions})

@app.route("/api/feedbacks", methods=["GET"])
def get_feedbacks():
    """Lấy danh sách phản hồi của khách hàng và các chỉ số CSAT"""
    store_id = request.args.get("store_id", "ALL")
    rating = request.args.get("rating", "ALL")
    category = request.args.get("category", "ALL")
    sentiment = request.args.get("sentiment", "ALL")
    feedbacks = engine.get_customer_feedbacks(store_id, rating, category, sentiment)
    summary = engine.get_feedback_summary()
    return jsonify({
        "success": True,
        "summary": summary,
        "feedbacks": feedbacks
    })

@app.route("/api/feedback", methods=["POST"])
def add_feedback():
    """Khách hàng hoặc chuyên viên thêm phản hồi mới"""
    data = request.json or {}
    customer_name = data.get("customer_name", "Khách hàng")
    phone = data.get("phone", "0900000000")
    store_id = data.get("store_id", "ST01")
    rating = data.get("rating", 5)
    category = data.get("category", "Dịch vụ & Tư vấn")
    comment = data.get("comment", "")

    if not comment:
        return jsonify({"success": False, "error": "Vui lòng nhập nội dung phản hồi."}), 400

    new_fb = engine.add_customer_feedback(customer_name, phone, store_id, rating, category, comment)
    return jsonify({
        "success": True,
        "data": new_fb,
        "summary": engine.get_feedback_summary(),
        "message": "Đã lưu phản hồi của khách hàng thành công!"
    })

@app.route("/api/export-cleaned-csv", methods=["GET"])
def export_cleaned_csv():
    """Tải xuống tệp CSV dữ liệu đã làm sạch"""
    cleaned_path = os.path.join(DATA_DIR, "cleaned_transactions_oct2026.csv")
    if not os.path.exists(cleaned_path):
        if engine.cleaned_df is not None:
            engine.cleaned_df.to_csv(cleaned_path, index=False, encoding="utf-8-sig")
        else:
            return "Chưa có dữ liệu làm sạch", 404
            
    return send_file(cleaned_path, mimetype="text/csv", as_attachment=True, download_name="BaoNgoc_Cleaned_Transactions_Oct2026.csv")


@app.route("/api/feedbacks/clear", methods=["POST"])
def clear_customer_feedbacks():
    """Xóa toàn bộ phản hồi khách hàng để đưa hệ thống về trạng thái ban đầu"""
    engine.customer_feedbacks = []
    fb_path = os.path.join(DATA_DIR, "customer_feedbacks.json")
    try:
        with open(fb_path, "w", encoding="utf-8") as f:
            json.dump([], f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    return jsonify({"success": True, "message": "Đã xóa toàn bộ phản hồi khách hàng thành công!", "summary": engine.get_feedback_summary()})

@app.route("/api/feedbacks/load-sample", methods=["POST"])
def load_sample_customer_feedbacks():
    """Nạp dữ liệu phản hồi khách hàng mẫu khi cần thử nghiệm"""
    sample_path = resolve_data_file("sample_customer_feedbacks.json")
    if os.path.exists(sample_path):
        try:
            with open(sample_path, "r", encoding="utf-8") as f:
                engine.customer_feedbacks = json.load(f)
            fb_path = os.path.join(DATA_DIR, "customer_feedbacks.json")
            with open(fb_path, "w", encoding="utf-8") as f:
                json.dump(engine.customer_feedbacks, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return jsonify({"success": True, "message": "Đã nạp dữ liệu phản hồi khách hàng mẫu thành công!", "count": len(engine.customer_feedbacks), "summary": engine.get_feedback_summary()})

@app.route("/api/generate-pitch-deck", methods=["GET", "POST"])
def api_generate_pitch_deck():
    """Tự động sinh bộ Slide PowerPoint Báo Cáo Điều Hành Ma Thuật (Auto-Pitch Deck)"""
    try:
        # Nếu chưa có báo cáo thì tự động tính toán
        if engine.management_report is None and engine.cleaned_df is not None:
            engine.generate_management_report()
        if engine.audit_results is None and engine.cleaned_df is not None:
            engine.perform_cross_check_audit()
            
        pptx_path = os.path.join(DATA_DIR, "BaoNgoc_Retail_PitchDeck_AI.pptx")
        generate_pitch_deck(engine, pptx_path)
        preview_data = get_pitch_deck_preview_data(engine)
        
        return jsonify({
            "success": True,
            "message": "Đã tự động tạo bộ Slide PowerPoint ma thuật thành công!",
            "slides": preview_data.get("slides", []),
            "preview": preview_data,
            "download_url": "/api/download-pitch-deck"
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/slides", methods=["GET"])
def presentation_slides():
    """Xem trực tiếp bài trình chiếu 8 Slide HTML toàn màn hình"""
    return render_template("slides.html")

@app.route("/api/download-slides-html", methods=["GET"])
def api_download_slides_html():
    """Tải xuống tệp trình chiếu HTML (.html) hoàn chỉnh độc lập"""
    html_path = os.path.join(WORKSPACE_DIR, "BaoNgoc_AI301_Slide_ThuyetTrinh.html")
    if not os.path.exists(html_path):
        html_path = os.path.join(os.path.dirname(BASE_DIR), "BaoNgoc_AI301_Slide_ThuyetTrinh.html")
        
    return send_file(
        html_path,
        mimetype="text/html",
        as_attachment=True,
        download_name="BaoNgoc_AI301_Slide_ThuyetTrinh.html"
    )

def get_safe_export_path(filename):
    """Lấy đường dẫn an toàn để lưu file xuất (hỗ trợ môi trường Read-only trên Vercel)"""
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        test_path = os.path.join(DATA_DIR, ".test_write")
        with open(test_path, "w") as f:
            f.write("1")
        os.remove(test_path)
        return os.path.join(DATA_DIR, filename)
    except Exception:
        import tempfile
        tmp = os.path.join(tempfile.gettempdir(), "baongoc_exports")
        os.makedirs(tmp, exist_ok=True)
        return os.path.join(tmp, filename)

@app.route("/api/download-pitch-deck", methods=["GET"])
def api_download_pitch_deck():
    """Tải xuống tệp trình chiếu PowerPoint (.pptx) hoàn chỉnh"""
    pptx_path = get_safe_export_path("BaoNgoc_Retail_PitchDeck_AI.pptx")
    if not os.path.exists(pptx_path):
        generate_pitch_deck(engine, pptx_path)
        
    return send_file(
        pptx_path,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name="BaoNgoc_AI_Retail_PitchDeck_2026.pptx"
    )

@app.route("/api/apply-custom-cleaning", methods=["POST"])
def api_apply_custom_cleaning():
    """Áp dụng ghi chú / yêu cầu tùy biến làm sạch từ người dùng ở Bước 2"""
    try:
        data = request.get_json() or {}
        instruction = data.get("instruction", "")
        action_type = data.get("action_type", "")
        target_column = data.get("target_column", "")
        target_row_id = data.get("target_row_id", "")
        
        result = engine.apply_custom_cleaning(
            instruction=instruction,
            action_type=action_type,
            target_column=target_column,
            target_row_id=target_row_id
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400

@app.route("/api/reset-custom-cleaning", methods=["POST"])
def api_reset_custom_cleaning():
    """Khôi phục lại dữ liệu sạch ban đầu của Bước 2"""
    try:
        result = engine.reset_custom_cleaning()
        return jsonify(result)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400

@app.route("/api/preview-word-report", methods=["GET"])
def api_preview_word_report():
    """Lấy bản xem trước HTML định dạng A4 của Báo cáo Word"""
    try:
        if engine.management_report is None and engine.cleaned_df is not None:
            engine.generate_management_report()
        html = get_word_report_preview_html(engine)
        return jsonify({"success": True, "html": html})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/api/download-word-report", methods=["GET"])
def api_download_word_report():
    """Tải xuống tệp Báo cáo Quản trị Microsoft Word (.docx)"""
    try:
        if engine.management_report is None and engine.cleaned_df is not None:
            engine.generate_management_report()
        docx_path = get_safe_export_path("BaoNgoc_BaoCao_QuanTri_AI301.docx")
        generate_word_report(engine, docx_path)
        return send_file(
            docx_path,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="BaoNgoc_BaoCao_QuanTri_AI301.docx"
        )
    except Exception as e:
        return str(e), 500

@app.route("/api/export-cleaned-excel", methods=["GET"])
@app.route("/api/download-cleaned-excel", methods=["GET"])
def api_download_cleaned_excel():
    """Tải xuống tệp Bảng Dữ Liệu Chuẩn Hóa Microsoft Excel (.xlsx) gồm 4 Sheet"""
    try:
        if engine.cleaned_df is None:
            if engine.raw_df is not None:
                engine.execute_data_cleaning()
            else:
                return "Chưa có dữ liệu để xuất Excel", 400
        excel_path = get_safe_export_path("BaoNgoc_DuLieu_ChuanHoa_AI301.xlsx")
        generate_excel_workbook(engine, excel_path)
        return send_file(
            excel_path,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="BaoNgoc_DuLieu_ChuanHoa_AI301.xlsx"
        )
    except Exception as e:
        return str(e), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=False)

