# update_template_empty_state.py
import re

html_path = r"c:\Users\ADMIN\OneDrive\Attachments\Thi cuối môn AI301\app\templates\index.html"
with open(html_path, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Update loadFeedbacks() KPI update logic
old_kpi_updates = """            // Update KPI Cards
            if (document.getElementById('fbKpiCsat')) document.getElementById('fbKpiCsat').innerText = sum.average_csat || '4.6';
            if (document.getElementById('fbKpiPosPct')) document.getElementById('fbKpiPosPct').innerText = `${sum.positive_pct || 85}%`;
            if (document.getElementById('fbKpiTotal')) document.getElementById('fbKpiTotal').innerText = sum.total_reviews || feedbackList.length;
            if (document.getElementById('fbKpiResolution')) document.getElementById('fbKpiResolution').innerText = `${sum.resolution_rate || 96.5}%`;
            if (document.getElementById('fbKpiRiskCount')) document.getElementById('fbKpiRiskCount').innerText = sum.risk_count || 0;
            if (document.getElementById('fbFilteredCount')) document.getElementById('fbFilteredCount').innerText = feedbackList.length;
            if (document.getElementById('sbBadgeFeedback')) document.getElementById('sbBadgeFeedback').innerText = `${sum.average_csat || 4.6}★`;"""

new_kpi_updates = """            // Update KPI Cards (Handles clean/empty customer data state)
            const hasReviews = sum.total_reviews > 0;
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
            }

            // Update Spotlight Banner on Step 3
            const spotlightBanner = document.getElementById('spotlightCsatBanner');
            if (spotlightBanner) {
                if (hasReviews) {
                    spotlightBanner.innerHTML = `
                        <div class="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                            <div class="flex items-start gap-4">
                                <div class="w-14 h-14 rounded-2xl bg-gradient-to-br from-amber-400 via-rose-500 to-pink-500 text-white flex flex-col items-center justify-center flex-shrink-0 shadow-md">
                                    <span class="text-xl font-black leading-none">${sum.average_csat}★</span>
                                    <span class="text-[9px] font-black uppercase tracking-wider mt-0.5">CSAT</span>
                                </div>
                                <div>
                                    <div class="flex flex-wrap items-center gap-2">
                                        <h3 class="font-black text-sm sm:text-base text-slate-900 dark:text-white flex items-center gap-1.5">
                                            <i class="fa-solid fa-comments text-amber-500"></i>
                                            Phản Hồi Khách Hàng &amp; Đo Lường CSAT Thực Tế
                                        </h3>
                                        <span class="px-2.5 py-0.5 rounded-full text-[10px] font-black bg-rose-100 text-rose-700 dark:bg-rose-950/60 dark:text-rose-300 border border-rose-200">
                                            ${sum.total_reviews} Đánh Giá Đa Kênh
                                        </span>
                                        <span class="px-2.5 py-0.5 rounded-full text-[10px] font-black bg-emerald-100 text-emerald-700 dark:bg-emerald-950/60 dark:text-emerald-300 border border-emerald-200">
                                            ${sum.positive_pct}% Hài Lòng
                                        </span>
                                    </div>
                                    <div class="mt-2.5 grid grid-cols-1 md:grid-cols-2 gap-2 text-xs">
                                        <div class="p-2.5 rounded-xl bg-white/90 dark:bg-slate-800/90 border border-rose-200 dark:border-slate-700 flex items-start gap-2 shadow-2xs">
                                            <i class="fa-solid fa-triangle-exclamation text-rose-500 mt-0.5 text-xs flex-shrink-0"></i>
                                            <div>
                                                <b class="text-rose-600 font-extrabold">ST09 (Bắc Từ Liêm):</b> 
                                                <span class="text-slate-600 dark:text-slate-300">Khách phản ánh móp lon sữa Meiji &amp; date cận (còn 2 tháng) &rarr; Tỷ lệ đổi trả <b class="text-rose-600">14.8%</b>.</span>
                                            </div>
                                        </div>
                                        <div class="p-2.5 rounded-xl bg-white/90 dark:bg-slate-800/90 border border-amber-200 dark:border-slate-700 flex items-start gap-2 shadow-2xs">
                                            <i class="fa-solid fa-box-open text-amber-500 mt-0.5 text-xs flex-shrink-0"></i>
                                            <div>
                                                <b class="text-amber-600 font-extrabold">ST18 (Cần Thơ):</b> 
                                                <span class="text-slate-600 dark:text-slate-300">Khách phàn nàn thiếu bỉm Merries/Moony size M &amp; L vào cuối tuần &rarr; Sụt giảm DT <b class="text-amber-600">-33.9% MoM</b>.</span>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="flex flex-wrap sm:flex-nowrap items-center gap-2.5 flex-shrink-0 self-start lg:self-center">
                                <button onclick="switchTab('feedback')" class="btn-pink-pill px-4 py-2.5 text-xs flex items-center gap-2 shadow-md hover:scale-105 transition cursor-pointer font-extrabold">
                                    <i class="fa-solid fa-arrow-right-to-bracket"></i>
                                    <span>Xem Toàn Bộ Đánh Giá &amp; Biểu Đồ &rarr;</span>
                                </button>
                                <button onclick="openFeedbackModal()" class="btn-outline-clean px-3.5 py-2.5 text-xs flex items-center gap-2 font-bold cursor-pointer hover:border-rose-400">
                                    <i class="fa-solid fa-plus text-rose-500"></i>
                                    <span>+ Gửi Phản Hồi Mới</span>
                                </button>
                            </div>
                        </div>`;
                } else {
                    spotlightBanner.innerHTML = `
                        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div class="flex items-center gap-3.5">
                                <div class="w-12 h-12 rounded-2xl bg-rose-100 text-rose-600 dark:bg-rose-950/60 dark:text-rose-400 flex items-center justify-center text-xl flex-shrink-0 shadow-xs">
                                    <i class="fa-regular fa-comment-dots"></i>
                                </div>
                                <div>
                                    <div class="flex items-center gap-2">
                                        <h3 class="font-extrabold text-sm sm:text-base text-slate-900 dark:text-white">
                                            Phản Hồi Khách Hàng (Customer Feedback)
                                        </h3>
                                        <span class="px-2 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 dark:bg-slate-800 text-slate-500">
                                            Chưa Có Dữ Liệu
                                        </span>
                                    </div>
                                    <p class="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
                                        Hệ thống đang ở trạng thái sạch ban đầu. Khách hàng hoặc quản trị viên có thể gửi phản hồi để bắt đầu đo lường CSAT.
                                    </p>
                                </div>
                            </div>
                            <div class="flex items-center gap-2.5 flex-shrink-0">
                                <button onclick="openFeedbackModal()" class="btn-pink-pill px-4 py-2 text-xs font-bold flex items-center gap-1.5 shadow-md">
                                    <i class="fa-solid fa-plus"></i>
                                    <span>+ Gửi Phản Hồi Đầu Tiên</span>
                                </button>
                                <button onclick="switchTab('feedback')" class="btn-outline-clean px-3.5 py-2 text-xs font-bold">
                                    <span>Xem Chi Tiết</span>
                                </button>
                            </div>
                        </div>`;
                }
            }"""

if old_kpi_updates in content:
    content = content.replace(old_kpi_updates, new_kpi_updates, 1)
    print("Updated loadFeedbacks() KPI logic successfully")
else:
    print("Warning: old_kpi_updates not found")

# 2. Update renderFeedbackCards for empty state
old_render_cards = """        if (!list || list.length === 0) {
            box.innerHTML = `<div class="col-span-full py-12 text-center text-slate-400 font-bold text-xs"><i class="fa-solid fa-inbox text-3xl mb-2 block"></i>Không tìm thấy phản hồi nào phù hợp với bộ lọc đã chọn.</div>`;
            return;
        }"""

new_render_cards = """        if (!list || list.length === 0) {
            box.innerHTML = `
                <div class="col-span-full os-card p-12 text-center flex flex-col items-center justify-center space-y-4">
                    <div class="w-16 h-16 rounded-3xl bg-rose-50 text-rose-500 dark:bg-rose-950/40 flex items-center justify-center text-3xl shadow-inner">
                        <i class="fa-regular fa-comments"></i>
                    </div>
                    <div>
                        <h4 class="font-black text-slate-800 dark:text-white text-base">Hệ Thống Chưa Có Dữ Liệu Khách Hàng Nào</h4>
                        <p class="text-xs text-slate-400 mt-1 max-w-md mx-auto">
                            Trang web hiện đang ở trạng thái sạch ban đầu, chưa ghi nhận phản hồi hoặc đánh giá nào từ khách hàng. Bạn có thể gửi phản hồi mới ngay bây giờ hoặc nạp dữ liệu mẫu khi cần thử nghiệm!
                        </p>
                    </div>
                    <div class="flex flex-wrap items-center justify-center gap-3 pt-2">
                        <button onclick="openFeedbackModal()" class="btn-pink-pill px-4 py-2 text-xs font-bold flex items-center gap-2 cursor-pointer shadow-md">
                            <i class="fa-solid fa-plus"></i>
                            <span>+ Gửi Đánh Giá Đầu Tiên</span>
                        </button>
                        <button onclick="loadSampleCustomerFeedbacks()" class="btn-outline-clean px-4 py-2 text-xs font-bold flex items-center gap-2 cursor-pointer">
                            <i class="fa-solid fa-wand-magic-sparkles text-amber-500"></i>
                            <span>Nạp Dữ Liệu Khách Hàng Mẫu</span>
                        </button>
                    </div>
                </div>`;
            return;
        }"""

if old_render_cards in content:
    content = content.replace(old_render_cards, new_render_cards, 1)
    print("Updated renderFeedbackCards() with rich Empty State successfully")
else:
    print("Warning: old_render_cards not found")

# 3. Add id="spotlightCsatBanner" to the Step 3 Spotlight Card
old_spotlight_banner = '<!-- ⭐ CSAT & CUSTOMER FEEDBACK SPOTLIGHT BANNER (Hiển thị trực tiếp ngay trên Dashboard chính Bước 3) -->\n                <div class="os-card p-5'
new_spotlight_banner = '<!-- ⭐ CSAT & CUSTOMER FEEDBACK SPOTLIGHT BANNER (Hiển thị trực tiếp ngay trên Dashboard chính Bước 3) -->\n                <div id="spotlightCsatBanner" class="os-card p-5'

if old_spotlight_banner in content:
    content = content.replace(old_spotlight_banner, new_spotlight_banner, 1)
    print("Added id='spotlightCsatBanner' successfully")
else:
    print("Warning: old_spotlight_banner not found")

# 4. Add reset & load-sample helper functions to JavaScript
helper_funcs = """
        // =====================================================================
        // EMPTY / SAMPLE CUSTOMER DATA MANAGEMENT
        // =====================================================================
        async function clearCustomerFeedbacks() {
            if (!confirm('Bạn có chắc muốn xóa toàn bộ dữ liệu phản hồi khách hàng để đưa hệ thống về trạng thái sạch ban đầu (chưa có dữ liệu)?')) return;
            try {
                const res = await (await fetch('/api/feedbacks/clear', { method: 'POST' })).json();
                if (res.success) {
                    showToast(res.message, 'success');
                    loadFeedbacks();
                }
            } catch (err) {
                console.error(err);
            }
        }

        async function loadSampleCustomerFeedbacks() {
            try {
                const res = await (await fetch('/api/feedbacks/load-sample', { method: 'POST' })).json();
                if (res.success) {
                    showToast(res.message, 'success');
                    loadFeedbacks();
                }
            } catch (err) {
                console.error(err);
            }
        }
"""

init_lang_target = "        function initLanguage()"
if init_lang_target in content:
    content = content.replace(init_lang_target, helper_funcs + "\n" + init_lang_target, 1)
    print("Added clear/load-sample JS helper functions successfully")
else:
    print("Warning: init_lang_target not found")

# 5. Add "Quản lý dữ liệu mẫu" buttons in the Feedback View Header
old_fb_view_header = """                            <button onclick="openFeedbackModal()" class="btn-pink-pill px-4 py-2.5 text-xs flex items-center gap-2 cursor-pointer shadow-md">
                                <i class="fa-solid fa-plus"></i>
                                <span>+ Gửi Phản Hồi Mới</span>
                            </button>"""

new_fb_view_header = """                            <button onclick="openFeedbackModal()" class="btn-pink-pill px-4 py-2.5 text-xs flex items-center gap-2 cursor-pointer shadow-md">
                                <i class="fa-solid fa-plus"></i>
                                <span>+ Gửi Phản Hồi Mới</span>
                            </button>
                            <button onclick="clearCustomerFeedbacks()" title="Đưa về trạng thái sạch ban đầu (Chưa có dữ liệu)" class="btn-outline-clean px-3 py-2.5 text-xs flex items-center gap-1.5 cursor-pointer text-slate-500 hover:text-rose-600">
                                <i class="fa-solid fa-trash-can"></i>
                                <span class="hidden sm:inline">Xóa Dữ Liệu</span>
                            </button>
                            <button onclick="loadSampleCustomerFeedbacks()" title="Nạp lại dữ liệu khách hàng mẫu thử nghiệm" class="btn-outline-clean px-3 py-2.5 text-xs flex items-center gap-1.5 cursor-pointer text-slate-500 hover:text-amber-600">
                                <i class="fa-solid fa-wand-magic-sparkles text-amber-500"></i>
                                <span class="hidden sm:inline">Nạp Mẫu</span>
                            </button>"""

if old_fb_view_header in content:
    content = content.replace(old_fb_view_header, new_fb_view_header, 1)
    print("Added action buttons to feedback view header successfully")
else:
    print("Warning: old_fb_view_header not found")

with open(html_path, "w", encoding="utf-8") as f:
    f.write(content)

print("index.html successfully updated for empty customer data state!")

