# patch_empty_ui.py
html_path = r"c:\Users\ADMIN\OneDrive\Attachments\Thi cuối môn AI301\app\templates\index.html"
with open(html_path, "r", encoding="utf-8") as f:
    content = f.read()

target_kpi = """                // Update KPI Cards
                if (document.getElementById('fbKpiCsat')) document.getElementById('fbKpiCsat').innerText = sum.average_csat || '4.6';
                if (document.getElementById('fbKpiPosPct')) document.getElementById('fbKpiPosPct').innerText = `${sum.positive_pct || 85}%`;
                if (document.getElementById('fbKpiTotal')) document.getElementById('fbKpiTotal').innerText = sum.total_reviews || feedbackList.length;
                if (document.getElementById('fbKpiResolution')) document.getElementById('fbKpiResolution').innerText = `${sum.resolution_rate || 96.5}%`;
                if (document.getElementById('fbKpiRiskCount')) document.getElementById('fbKpiRiskCount').innerText = sum.risk_count || 0;
                if (document.getElementById('fbFilteredCount')) document.getElementById('fbFilteredCount').innerText = feedbackList.length;
                if (document.getElementById('sbBadgeFeedback')) document.getElementById('sbBadgeFeedback').innerText = `${sum.average_csat || 4.6}★`;"""

replacement_kpi = """                // Update KPI Cards (Supports Empty Customer Data State)
                const hasReviews = (sum.total_reviews !== undefined && sum.total_reviews > 0);
                if (document.getElementById('fbKpiCsat')) {
                    document.getElementById('fbKpiCsat').innerText = hasReviews ? `${sum.average_csat}★` : '--';
                }
                if (document.getElementById('fbKpiPosPct')) {
                    document.getElementById('fbKpiPosPct').innerText = hasReviews ? `${sum.positive_pct}%` : '0%';
                }
                if (document.getElementById('fbKpiTotal')) {
                    document.getElementById('fbKpiTotal').innerText = sum.total_reviews || 0;
                }
                if (document.getElementById('fbKpiResolution')) {
                    document.getElementById('fbKpiResolution').innerText = hasReviews ? `${sum.resolution_rate}%` : '0%';
                }
                if (document.getElementById('fbKpiRiskCount')) {
                    document.getElementById('fbKpiRiskCount').innerText = sum.risk_count || 0;
                }
                if (document.getElementById('fbFilteredCount')) {
                    document.getElementById('fbFilteredCount').innerText = feedbackList.length;
                }
                if (document.getElementById('sbBadgeFeedback')) {
                    document.getElementById('sbBadgeFeedback').innerText = hasReviews ? `${sum.average_csat}★` : '0★ Mới';
                }"""

target_cards = """            if (!list || list.length === 0) {
                box.innerHTML = `<div class="col-span-full py-12 text-center text-slate-400 font-bold text-xs"><i class="fa-solid fa-inbox text-3xl mb-2 block"></i>Không tìm thấy phản hồi nào phù hợp với bộ lọc đã chọn.</div>`;
                return;
            }"""

replacement_cards = """            if (!list || list.length === 0) {
                box.innerHTML = `
                    <div class="col-span-full os-card p-12 text-center flex flex-col items-center justify-center space-y-4">
                        <div class="w-16 h-16 rounded-3xl bg-rose-50 text-rose-500 dark:bg-rose-950/40 flex items-center justify-center text-3xl shadow-inner">
                            <i class="fa-regular fa-comments"></i>
                        </div>
                        <div>
                            <h4 class="font-black text-slate-800 dark:text-white text-base">Hệ Thống Chưa Có Dữ Liệu Khách Hàng</h4>
                            <p class="text-xs text-slate-400 mt-1 max-w-md mx-auto">
                                Trang web đang ở trạng thái sạch ban đầu, chưa ghi nhận phản hồi hay khiếu nại nào từ khách hàng. Bạn có thể gửi phản hồi mới ngay bây giờ hoặc nạp dữ liệu mẫu thử nghiệm!
                            </p>
                        </div>
                        <div class="flex flex-wrap items-center justify-center gap-3 pt-2">
                            <button onclick="openFeedbackModal()" class="btn-pink-pill px-4 py-2 text-xs font-bold flex items-center gap-2 cursor-pointer shadow-md">
                                <i class="fa-solid fa-plus"></i>
                                <span>+ Gửi Phản Hồi Đầu Tiên</span>
                            </button>
                            <button onclick="loadSampleCustomerFeedbacks()" class="btn-outline-clean px-4 py-2 text-xs font-bold flex items-center gap-2 cursor-pointer">
                                <i class="fa-solid fa-wand-magic-sparkles text-amber-500"></i>
                                <span>Nạp Dữ Liệu Khách Hàng Mẫu</span>
                            </button>
                        </div>
                    </div>`;
                return;
            }"""

if target_kpi in content:
    content = content.replace(target_kpi, replacement_kpi, 1)
    print("KPI logic replaced successfully")
else:
    print("Warning: target_kpi not found")

if target_cards in content:
    content = content.replace(target_cards, replacement_cards, 1)
    print("Cards empty state replaced successfully")
else:
    print("Warning: target_cards not found")

with open(html_path, "w", encoding="utf-8") as f:
    f.write(content)

print("patch_empty_ui.py executed successfully!")

