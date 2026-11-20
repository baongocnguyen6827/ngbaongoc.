# set_empty_customer_data.py
import os
import json

base_dir = r"c:\Users\ADMIN\OneDrive\Attachments\Thi cuối môn AI301"
data_dir = os.path.join(base_dir, "app", "data")
fb_path = os.path.join(data_dir, "customer_feedbacks.json")
sample_path = os.path.join(data_dir, "sample_customer_feedbacks.json")

# 1. Backup current 21 feedbacks into sample_customer_feedbacks.json
if os.path.exists(fb_path):
    with open(fb_path, "r", encoding="utf-8") as f:
        current_data = json.load(f)
    if len(current_data) > 0:
        with open(sample_path, "w", encoding="utf-8") as f:
            json.dump(current_data, f, ensure_ascii=False, indent=2)
        print(f"Backed up {len(current_data)} sample feedbacks to {sample_path}")

# 2. Clear customer_feedbacks.json to empty list []
with open(fb_path, "w", encoding="utf-8") as f:
    json.dump([], f, ensure_ascii=False, indent=2)
print("customer_feedbacks.json is now EMPTY: []")

# 3. Update app/engine.py get_feedback_summary() to return 0.0 when empty
engine_path = os.path.join(base_dir, "app", "engine.py")
with open(engine_path, "r", encoding="utf-8") as f:
    engine_code = f.read()

old_summary_zero = """        if total == 0:
            return {
                "total_reviews": 0,
                "average_csat": 5.0,
                "positive_pct": 100.0,
                "neutral_pct": 0.0,
                "negative_pct": 0.0,
                "resolution_rate": 100.0,
                "rating_counts": {5: 0, 4: 0, 3: 0, 2: 0, 1: 0},
                "category_counts": {},
                "risk_count": 0
            }"""

new_summary_zero = """        if total == 0:
            return {
                "total_reviews": 0,
                "average_csat": 0.0,
                "positive_pct": 0.0,
                "neutral_pct": 0.0,
                "negative_pct": 0.0,
                "resolution_rate": 0.0,
                "rating_counts": {5: 0, 4: 0, 3: 0, 2: 0, 1: 0},
                "category_counts": {},
                "risk_count": 0
            }"""

if old_summary_zero in engine_code:
    engine_code = engine_code.replace(old_summary_zero, new_summary_zero, 1)
    with open(engine_path, "w", encoding="utf-8") as f:
        f.write(engine_code)
    print("Updated engine.py get_feedback_summary() for 0 total reviews")
else:
    print("Note: old_summary_zero not found or already updated")

# 4. Update app/main.py to add /api/feedbacks/clear and /api/feedbacks/load-sample endpoints
main_path = os.path.join(base_dir, "app", "main.py")
with open(main_path, "r", encoding="utf-8") as f:
    main_code = f.read()

endpoints_code = """
@app.route("/api/feedbacks/clear", methods=["POST"])
def clear_customer_feedbacks():
    \"\"\"Xóa toàn bộ phản hồi khách hàng để đưa hệ thống về trạng thái ban đầu\"\"\"
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
    \"\"\"Nạp dữ liệu phản hồi khách hàng mẫu khi cần thử nghiệm\"\"\"
    sample_path = os.path.join(DATA_DIR, "sample_customer_feedbacks.json")
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
"""

if "/api/feedbacks/clear" not in main_code:
    # insert before "if __name__ == '__main__':"
    if "if __name__ ==" in main_code:
        main_code = main_code.replace("if __name__ ==", endpoints_code + "\nif __name__ ==")
    else:
        main_code += endpoints_code
    with open(main_path, "w", encoding="utf-8") as f:
        f.write(main_code)
    print("Added clear and load-sample endpoints to app/main.py")
else:
    print("Endpoints already present in app/main.py")

print("All backend changes for empty customer data completed!")

