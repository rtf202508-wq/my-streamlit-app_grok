import os
import sys
import glob
import re
import datetime
import base64
import time
from io import BytesIO
import openpyxl
from PIL import Image, ImageDraw
from openai import OpenAI
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import streamlit as st
import tkinter as tk
from tkinter import filedialog
from itertools import zip_longest

# ---------------------------------------------------------
# 0. 페이지 기본 설정 및 OS 폴더 다이얼로그 헬퍼
# ---------------------------------------------------------
st.set_page_config(
    page_title="이미지 QA 자동 검수 시스템 v9.0 Grok",
    page_icon="🔍",
    layout="wide"
)

# [Grok API 키 설정]
raw_api_key = os.environ.get("XAI_API_KEY", "")
api_key = "".join(c for c in raw_api_key if 32 < ord(c) < 127).strip()

if not api_key:
    st.error("오류: 환경변수 XAI_API_KEY가 설정되지 않았습니다. console.x.ai에서 키를 발급받아 설정해주세요.")
    st.stop()

client = OpenAI(
    api_key=api_key,
    base_url="https://api.x.ai/v1"
)

os.makedirs("temp_uploads", exist_ok=True)
os.makedirs("annotated_output", exist_ok=True)
os.makedirs("history", exist_ok=True)

def open_folder_dialog(title="폴더 선택"):
    """OS 네이티브 폴더 선택창 호출"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder_path = filedialog.askdirectory(title=title)
    root.destroy()
    return folder_path

def get_images_in_folder_recursive(folder_path):
    """하위 폴더 포함 모든 이미지 확장자 검색 및 디버깅 출력"""
    img_extensions = ('.png', '.jpg', '.jpeg', '.webp', '.bmp', '.psd', '.tiff', '.tif', '.gif')
    files = []
    if folder_path:
        clean_path = folder_path.strip().strip('"').strip("'")
        if os.path.exists(clean_path):
            for root, dirs, filenames in os.walk(clean_path):
                for f in filenames:
                    if f.lower().endswith(img_extensions):
                        files.append(os.path.join(root, f))
            print(f"[DEBUG] 탐색 경로: {clean_path} -> 찾은 이미지 수: {len(files)}개")
        else:
            print(f"[DEBUG] 경로가 존재하지 않습니다: {clean_path}")
    return sorted(files)

def save_uploaded_file(uploaded_file):
    temp_path = os.path.join("temp_uploads", uploaded_file.name)
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return temp_path

def image_to_base64_data_uri(image_path_or_pil, max_size=(2048, 2048), quality=85):
    """이미지를 JPEG base64 data URI로 변환 (Grok 지원 포맷)"""
    try:
        if isinstance(image_path_or_pil, str):
            img = Image.open(image_path_or_pil)
        else:
            img = image_path_or_pil
        
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        
        # 크기 제한 (토큰/비용 절약)
        img.thumbnail(max_size, Image.Resampling.LANCZOS)
        
        buffered = BytesIO()
        img.save(buffered, format="JPEG", quality=quality, optimize=True)
        b64 = base64.b64encode(buffered.getvalue()).decode("utf-8")
        return f"data:image/jpeg;base64,{b64}"
    except Exception as e:
        print(f"[WARN] 이미지 변환 실패: {e}")
        return None

def draw_diff_boxes_on_image(image_path, report_text, output_path):
    """Grok가 반환한 BOX 좌표를 바탕으로 결과물 이미지에 빨간색 박스 드로잉"""
    try:
        img = Image.open(image_path).convert("RGB")
    except Exception:
        img = Image.new("RGB", (1000, 1000), color=(255, 255, 255))
    
    draw = ImageDraw.Draw(img)
    w, h = img.size

    box_pattern = re.findall(r'BOX\s*:\s*\[\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\]', report_text, re.IGNORECASE)
    
    drawn_count = 0
    for b in box_pattern:
        y1, x1, y2, x2 = map(int, b)
        if max(y1, x1, y2, x2) <= 1000:
            ax1 = int(x1 / 1000.0 * w)
            ay1 = int(y1 / 1000.0 * h)
            ax2 = int(x2 / 1000.0 * w)
            ay2 = int(y2 / 1000.0 * h)
        else:
            ax1, ay1, ax2, ay2 = x1, y1, x2, y2
        
        line_width = max(3, int(min(w, h) * 0.008))
        draw.rectangle([ax1, ay1, ax2, ay2], outline="red", width=line_width)
        drawn_count += 1
    
    img.save(output_path)
    return output_path, drawn_count

def call_grok_with_retry(prompt_text, image_data_uris, model="grok-4.3", max_retries=3):
    """Grok API 호출 + 재시도"""
    content = [{"type": "text", "text": prompt_text}]
    for uri in image_data_uris:
        if uri:
            content.append({
                "type": "image_url",
                "image_url": {"url": uri, "detail": "high"}
            })
    
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": content}],
                temperature=0.1,
                max_tokens=4096
            )
            return response.choices[0].message.content
        except Exception as e:
            if attempt == max_retries - 1:
                raise e
            wait = 2 ** attempt
            print(f"[RETRY] 시도 {attempt+1} 실패, {wait}초 후 재시도: {e}")
            time.sleep(wait)
    return None

# ---------------------------------------------------------
# 세션 상태 초기화 및 동기화 콜백 함수 정의
# ---------------------------------------------------------
if "orig_dir_path" not in st.session_state: st.session_state["orig_dir_path"] = ""
if "orig_dir_input_v89" not in st.session_state: st.session_state["orig_dir_input_v89"] = st.session_state["orig_dir_path"]

if "res_dir_path" not in st.session_state: st.session_state["res_dir_path"] = ""
if "res_dir_input_v89" not in st.session_state: st.session_state["res_dir_input_v89"] = st.session_state["res_dir_path"]

if "guide_dir_path" not in st.session_state: st.session_state["guide_dir_path"] = ""
if "guide_dir_input_v89" not in st.session_state: st.session_state["guide_dir_input_v89"] = st.session_state["guide_dir_path"]

if "req_dir_path" not in st.session_state: st.session_state["req_dir_path"] = ""
if "req_dir_input_v89" not in st.session_state: st.session_state["req_dir_input_v89"] = st.session_state["req_dir_path"]

def update_orig_path():
    st.session_state["orig_dir_path"] = st.session_state["orig_dir_input_v89"]

def update_res_path():
    st.session_state["res_dir_path"] = st.session_state["res_dir_input_v89"]

def update_guide_path():
    st.session_state["guide_dir_path"] = st.session_state["guide_dir_input_v89"]

def update_req_path():
    st.session_state["req_dir_path"] = st.session_state["req_dir_input_v89"]

# ---------------------------------------------------------
# 2. 사이드바 - 설정 영역 (v9.0 Grok)
# ---------------------------------------------------------
st.sidebar.title("🛠️ QA 검수 설정 (v9.0 Grok)")

grok_model = st.sidebar.selectbox(
    "Grok 모델 선택",
    ["grok-4.3", "grok-4.20", "grok-4.1-fast", "grok-4"],
    index=0,
    help="비전 지원 모델. grok-4.3 권장"
)

selected_type = st.sidebar.selectbox(
    "1. 검수 대상 이미지 유형을 선택하세요:",
    ["캐리어 페이지", "썸네일", "Infovisual", "상세페이지", "로컬라이징", "소셜 이미지", "스펙표", "기타"],
    key="sel_type_v89"
)

carrier_info = ""
if selected_type == "캐리어 페이지":
    client_name = st.sidebar.text_input("고객사명 입력 (예: SEJ)", value="SEJ", key="client_name_v89")
    carrier_choice = st.sidebar.selectbox("통신사 선택", ["라쿠텐", "소프트뱅크", "DCM OLS", "DCM corp", "KDDI", "기타"], key="carrier_choice_v89")
    device_choice = st.sidebar.radio("매체 환경", ["PC", "MO"], key="device_choice_v89")
    carrier_info = f"\n- 고객사: {client_name}\n- 통신사: {carrier_choice}\n- 디바이스: {device_choice}"

# --- [그룹 A] 원본 이미지 ---
st.sidebar.markdown("---")
st.sidebar.subheader("🖼️ [그룹 A] 원본 이미지")
chk_orig_folder = st.sidebar.checkbox("폴더 선택 (하위폴더 포함)", key="chk_orig_f")
orig_folder_paths = []
if chk_orig_folder:
    col_orig_1, col_orig_2 = st.sidebar.columns([3, 1])
    with col_orig_2:
        if st.button("📁 폴더 열기", key="btn_open_orig"):
            sel_dir = open_folder_dialog("원본 이미지 폴더 선택")
            if sel_dir:
                st.session_state["orig_dir_path"] = sel_dir
                st.session_state["orig_dir_input_v89"] = sel_dir
                st.rerun()
    
    col_orig_1.text_input(
        "원본 폴더 경로", 
        key="orig_dir_input_v89", 
        on_change=update_orig_path
    )

    orig_folder_paths = get_images_in_folder_recursive(st.session_state["orig_dir_path"])
    st.sidebar.caption(f"📁 검색된 이미지 수: {len(orig_folder_paths)}개")
    if len(orig_folder_paths) == 0 and st.session_state["orig_dir_path"]:
        st.sidebar.warning("경로에 이미지(.jpg, .png, .psd 등)가 없거나 경로를 다시 확인해주세요.")

chk_orig_upload = st.sidebar.checkbox("파일 직접 업로드 (다중 선택)", key="chk_orig_u")
orig_uploads = []
if chk_orig_upload:
    orig_uploads = st.sidebar.file_uploader(
        "원본 이미지 다중 업로드",
        type=['png', 'jpg', 'jpeg', 'webp', 'bmp', 'psd', 'tiff', 'tif', 'gif'],
        accept_multiple_files=True,
        key="orig_up_v89"
    )

# --- [그룹 B] 결과물 이미지 ---
st.sidebar.markdown("---")
st.sidebar.subheader("🎨 [그룹 B] 결과물 이미지")
chk_res_folder = st.sidebar.checkbox("폴더 선택 (하위폴더 포함)", key="chk_res_f")
res_folder_paths = []
if chk_res_folder:
    col_res_1, col_res_2 = st.sidebar.columns([3, 1])
    with col_res_2:
        if st.button("📁 폴더 열기", key="btn_open_res"):
            sel_dir = open_folder_dialog("결과물 이미지 폴더 선택")
            if sel_dir:
                st.session_state["res_dir_path"] = sel_dir
                st.session_state["res_dir_input_v89"] = sel_dir
                st.rerun()
    
    col_res_1.text_input(
        "결과물 폴더 경로", 
        key="res_dir_input_v89", 
        on_change=update_res_path
    )

    res_folder_paths = get_images_in_folder_recursive(st.session_state["res_dir_path"])
    st.sidebar.caption(f"📁 검색된 이미지 수: {len(res_folder_paths)}개")
    if len(res_folder_paths) == 0 and st.session_state["res_dir_path"]:
        st.sidebar.warning("경로에 이미지(.jpg, .png, .psd 등)가 없거나 경로를 다시 확인해주세요.")

chk_res_upload = st.sidebar.checkbox("파일 직접 업로드 (다중 선택)", key="chk_res_u")
res_uploads = []
if chk_res_upload:
    res_uploads = st.sidebar.file_uploader(
        "결과물 이미지 다중 업로드",
        type=['png', 'jpg', 'jpeg', 'webp', 'bmp', 'psd', 'tiff', 'tif', 'gif'],
        accept_multiple_files=True,
        key="res_up_v89"
    )

# --- [그룹 Guide] 가이드/라인 이미지 (선택사항) ---
st.sidebar.markdown("---")
st.sidebar.subheader("📐 [그룹 Guide] 가이드/라인 이미지 (선택사항)")
chk_guide_folder = st.sidebar.checkbox("가이드 폴더 선택 (하위폴더 포함)", key="chk_guide_f")
guide_folder_paths = []
if chk_guide_folder:
    col_guide_1, col_guide_2 = st.sidebar.columns([3, 1])
    with col_guide_2:
        if st.button("📁 폴더 열기", key="btn_open_guide"):
            sel_dir = open_folder_dialog("가이드 이미지 폴더 선택")
            if sel_dir:
                st.session_state["guide_dir_path"] = sel_dir
                st.session_state["guide_dir_input_v89"] = sel_dir
                st.rerun()
    
    col_guide_1.text_input(
        "가이드 폴더 경로", 
        key="guide_dir_input_v89", 
        on_change=update_guide_path
    )

    guide_folder_paths = get_images_in_folder_recursive(st.session_state["guide_dir_path"])
    st.sidebar.caption(f"📁 검색된 이미지 수: {len(guide_folder_paths)}개")

chk_guide_upload = st.sidebar.checkbox("가이드 파일 직접 업로드 (다중 선택)", key="chk_guide_u")
guide_uploads = []
if chk_guide_upload:
    guide_uploads = st.sidebar.file_uploader(
        "가이드/라인 이미지 다중 업로드",
        type=['png', 'jpg', 'jpeg', 'webp', 'bmp', 'psd', 'tiff', 'tif', 'gif'],
        accept_multiple_files=True,
        key="guide_up_v89"
    )

# --- [그룹 C] 고객 요청사항 (선택사항) ---
st.sidebar.markdown("---")
st.sidebar.subheader("💬 [그룹 C] 고객 요청사항 (선택사항)")
chk_req_manual = st.sidebar.checkbox("수기 기재 (직접 작성)", key="chk_req_man")
manual_txt = ""
if chk_req_manual:
    manual_txt = st.sidebar.text_area("수기 기재 내용", placeholder="요청사항을 직접 입력하세요...", key="req_manual_txt_v89")

chk_req_folder = st.sidebar.checkbox("폴더 선택 (폴더 내 문서/이미지 파싱)", key="chk_req_f")
req_folder_texts = []
req_folder_images = []
if chk_req_folder:
    col_req_1, col_req_2 = st.sidebar.columns([3, 1])
    with col_req_2:
        if st.button("📁 폴더 열기", key="btn_open_req"):
            sel_dir = open_folder_dialog("요청사항 폴더 선택")
            if sel_dir:
                st.session_state["req_dir_path"] = sel_dir
                st.session_state["req_dir_input_v89"] = sel_dir
                st.rerun()
    
    col_req_1.text_input(
        "요청사항 폴더 경로", 
        key="req_dir_input_v89", 
        on_change=update_req_path
    )

    r_dir = st.session_state["req_dir_path"].strip().strip('"').strip("'")
    img_exts = ('.png', '.jpg', '.jpeg', '.webp', '.bmp', '.psd', '.tiff', '.tif', '.gif')
    if os.path.exists(r_dir):
        for root, dirs, files in os.walk(r_dir):
            for f in files:
                full_p = os.path.join(root, f)
                ext = os.path.splitext(f)[1].lower()
                if ext in img_exts:
                    req_folder_images.append(full_p)
                elif ext == '.docx':
                    doc_r = Document(full_p)
                    req_folder_texts.append("\n".join([p.text for p in doc_r.paragraphs if p.text.strip()]))
                elif ext == '.pptx':
                    prs_r = Presentation(full_p)
                    texts = []
                    for slide in prs_r.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame and shape.text.strip():
                                texts.append(shape.text.strip())
                    req_folder_texts.append("\n".join(texts))
                elif ext == '.txt':
                    with open(full_p, 'r', encoding='utf-8') as tf:
                        req_folder_texts.append(tf.read().strip())

chk_req_upload = st.sidebar.checkbox("파일 직접 업로드 (다중 선택)", key="chk_req_u")
req_file_uploads = []
if chk_req_upload:
    req_file_uploads = st.sidebar.file_uploader(
        "요청사항 파일 업로드 (txt, docx, pptx, 이미지 등 / 다중 선택)",
        type=['txt', 'docx', 'pptx', 'png', 'jpg', 'jpeg', 'webp', 'bmp', 'psd', 'tiff', 'tif', 'gif'],
        accept_multiple_files=True,
        key="req_up_v89"
    )

# ---------------------------------------------------------
# 3. 메인 로직 - 전체 파일 루프 처리 및 리포트/마킹 적용 (v9.0 Grok)
# ---------------------------------------------------------
st.title("🔍 통합 이미지 QA 자동 검수 시스템 v9.0 Grok")
st.caption("모든 이미지 파일 1:1 전수 검수 및 레드박스 마킹 지원 (Grok API)")

if st.button("🚀 QA 검수 및 리포트 생성 시작", type="primary", key="btn_start_qa_v89"):
    all_orig_paths = list(orig_folder_paths) + [save_uploaded_file(u) for u in orig_uploads]
    all_res_paths = list(res_folder_paths) + [save_uploaded_file(u) for u in res_uploads]

    if not all_orig_paths or not all_res_paths:
        st.error("오류: '원본 이미지'와 '결과물 이미지'가 최소 1장 이상 선택/업로드되어야 합니다.")
        st.stop()

    with st.spinner(f"Grok API ({grok_model})가 모든 이미지 쌍에 대해 정밀 분석을 진행 중입니다..."):
        try:
            all_guide_paths = list(guide_folder_paths) + [save_uploaded_file(u) for u in guide_uploads]

            req_texts = []
            req_image_paths = list(req_folder_images)

            if manual_txt.strip():
                req_texts.append(manual_txt.strip())
            if req_folder_texts:
                req_texts.extend(req_folder_texts)

            if req_file_uploads:
                for up in req_file_uploads:
                    fpath = save_uploaded_file(up)
                    ext = os.path.splitext(up.name)[1].lower()
                    img_exts = ['.png', '.jpg', '.jpeg', '.webp', '.bmp', '.psd', '.tiff', '.tif', '.gif']
                    if ext in img_exts:
                        req_image_paths.append(fpath)
                    elif ext == '.docx':
                        doc_req = Document(fpath)
                        req_texts.append("\n".join([p.text for p in doc_req.paragraphs if p.text.strip()]))
                    elif ext == '.pptx':
                        prs_req = Presentation(fpath)
                        texts = []
                        for slide in prs_req.slides:
                            for shape in slide.shapes:
                                if shape.has_text_frame and shape.text.strip():
                                    texts.append(shape.text.strip())
                        req_texts.append("\n".join(texts))
                    elif ext == '.txt':
                        with open(fpath, "r", encoding="utf-8") as tf:
                            req_texts.append(tf.read().strip())

            user_requests = "\n\n".join(req_texts) if req_texts else "특이 요청사항 없음"

            now = datetime.datetime.now()
            date_str = now.strftime("%Y%m%d")
            base_time_ver = now.strftime("%Y%m%d_%H%M_01")
            history_base_dir = "history"
            date_folder = os.path.join(history_base_dir, date_str)
            sub_folder = os.path.join(date_folder, base_time_ver)
            os.makedirs(sub_folder, exist_ok=True)

            pair_analysis_results = []
            overall_fail = False

            for idx, (orig_p, res_p) in enumerate(zip_longest(all_orig_paths, all_res_paths, fillvalue=None)):
                if not orig_p or not res_p:
                    continue

                # 이미지 data URI 준비
                image_uris = []
                for p in [orig_p, res_p]:
                    uri = image_to_base64_data_uri(p)
                    if uri:
                        image_uris.append(uri)
                
                for gp in all_guide_paths:
                    uri = image_to_base64_data_uri(gp)
                    if uri:
                        image_uris.append(uri)
                
                for rp_img in req_image_paths:
                    uri = image_to_base64_data_uri(rp_img)
                    if uri:
                        image_uris.append(uri)

                prompt = f"""
너는 최고 수준의 전문 이미지 QA(품질 검수) 엔지니어이자 시각 디자인/로컬라이징 검수 전문가야.
제공된 이미지 쌍 #{idx+1} ({os.path.basename(orig_p)} vs {os.path.basename(res_p)})을 검수해줘.

■ 기본 작업 정보
- 파일명: 원본({os.path.basename(orig_p)}) / 결과물({os.path.basename(res_p)})
- 검수 유형: {selected_type} {carrier_info}

■ 고객 요청 사항:
{user_requests}

■ 핵심 검수 가이드라인:
1. **1픽셀 단위 및 레이아웃/위치 정밀 검수**: 차이점이 발견되면 **결과물 이미지 기준**으로 박스 좌표를 `BOX: [ymin, xmin, ymax, xmax]` 형식으로 출력하시오 (0~1000 스케일).
2. **로컬라이징 및 일어 텍스트 검수**: 오탈자, 줄바꿈 부자연스러움, 텍스트 잘림 등을 전수 검토하시오.
3. 반드시 **[통과]** 또는 **[미통과]**를 명시하시오.
"""

                report_text = call_grok_with_retry(prompt, image_uris, model=grok_model)

                if not report_text:
                    report_text = "[오류] Grok API 응답 없음"

                is_fail = ("미통과" in report_text) or ("FAIL" in report_text.upper()) or ("불합격" in report_text)
                if is_fail:
                    overall_fail = True
                status_label = "미통과" if is_fail else "통과"

                annotated_img_filename = f"annotated_pair_{idx+1}_{status_label}.png"
                annotated_img_path = os.path.join(sub_folder, annotated_img_filename)
                annotated_res_path, box_count = draw_diff_boxes_on_image(res_p, report_text, annotated_img_path)

                pair_analysis_results.append({
                    "index": idx + 1,
                    "orig_path": orig_p,
                    "res_path": res_p,
                    "annotated_path": annotated_res_path,
                    "status_label": status_label,
                    "box_count": box_count,
                    "report_text": report_text
                })

            final_status_label = "미통과" if overall_fail else "통과"
            full_ver_str = f"{base_time_ver}_{final_status_label}"

            # 1) Word 리포트 (.docx)
            doc_filename = os.path.join(sub_folder, f"{full_ver_str}.docx")
            doc = Document()
            doc.add_heading(f"QA 검수 최종 리포트 ({selected_type}) - [{final_status_label}]", level=0)
            doc.add_paragraph(f"검수 일시: {now.strftime('%Y-%m-%d %H:%M:%S')}")
            doc.add_paragraph(f"종합 판정: {final_status_label}")
            doc.add_paragraph(f"작업 유형: {selected_type} {carrier_info}")
            doc.add_paragraph(f"사용 모델: {grok_model}")
            doc.add_paragraph(f"고객 요청사항:\n{user_requests}")
            
            for item in pair_analysis_results:
                doc.add_heading(f"[파일 Pair #{item['index']}] 판정: {item['status_label']} (마킹 박스: {item['box_count']}개)", level=1)
                doc.add_paragraph(f"원본 파일: {os.path.basename(item['orig_path'])}\n결과물 파일: {os.path.basename(item['res_path'])}")
                for paragraph in item['report_text'].split("\n\n"):
                    doc.add_paragraph(paragraph)
            doc.save(doc_filename)

            # 2) Excel 리포트 (.xlsx)
            xlsx_filename = os.path.join(sub_folder, f"{full_ver_str}.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "QA 검수 결과"
            ws.append(["구분", "상세 내용"])
            ws.append(["검수 일시", now.strftime('%Y-%m-%d %H:%M:%S')])
            ws.append(["종합 판정", final_status_label])
            ws.append(["검수 유형", selected_type])
            ws.append(["사용 모델", grok_model])
            ws.append(["고객 요청사항", user_requests])
            ws.append(["---", "---"])
            for item in pair_analysis_results:
                ws.append([f"Pair #{item['index']} ({os.path.basename(item['res_path'])}) 판정", item['status_label']])
                ws.append([f"Pair #{item['index']} 상세", item['report_text']])
            wb.save(xlsx_filename)

            # 3) PowerPoint 리포트 (.pptx)
            ppt_filename = os.path.join(sub_folder, f"{full_ver_str}.pptx")
            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

            for item in pair_analysis_results:
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.333), Inches(0.6))
                tf = txBox.text_frame
                p = tf.paragraphs[0]
                p.text = f"QA 검수 리포트 Pair #{item['index']} [{item['status_label']}] - {os.path.basename(item['res_path'])}"
                p.font.bold = True
                p.font.size = Pt(16)

                img_y, img_width = Inches(0.9), Inches(4.2)
                try: slide.shapes.add_picture(item['orig_path'], Inches(0.5), img_y, width=img_width)
                except: pass
                try: slide.shapes.add_picture(item['annotated_path'], Inches(5.0), img_y, width=img_width)
                except: pass

                line_color = RGBColor(255, 0, 0) if item['status_label'] == "미통과" else RGBColor(0, 180, 0)
                shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.0), Inches(2.2), Inches(4.2), Inches(0.04))
                shape.fill.solid()
                shape.fill.fore_color.rgb = line_color
                shape.line.color.rgb = line_color

                result_box = slide.shapes.add_textbox(Inches(9.4), Inches(0.9), Inches(3.4), Inches(6.0))
                tf_res = result_box.text_frame
                tf_res.word_wrap = True

                for line in item['report_text'].split("\n"):
                    if line.strip():
                        p_res = tf_res.add_paragraph()
                        p_res.text = line.strip()
                        p_res.font.size = Pt(9.0)
                        if "통과" in line or "결과" in line or "미통과" in line:
                            p_res.font.bold = True
                            p_res.font.size = Pt(10.0)

            prs.save(ppt_filename)

            st.success(f"✅ 전체 파일 QA 검수 완료! [종합 판정: {final_status_label} | 총 검수 파일 쌍: {len(pair_analysis_results)}개 | 모델: {grok_model}]")

            for item in pair_analysis_results:
                with st.expander(f"📁 파일 Pair #{item['index']} : {os.path.basename(item['res_path'])} [{item['status_label']}] (마킹 박스: {item['box_count']}개)", expanded=True):
                    col1, col2 = st.columns(2)
                    with col1:
                        st.subheader(f"원본 이미지: {os.path.basename(item['orig_path'])}")
                        try: st.image(item['orig_path'], use_container_width=True)
                        except: st.info("미리보기 미지원 형식 파일")
                    with col2:
                        st.subheader(f"결과물 이미지 (박스 {item['box_count']}곳 마킹)")
                        try: st.image(item['annotated_path'], use_container_width=True)
                        except: st.info("미리보기 미지원 형식 파일")
                    st.markdown("---")
                    st.markdown(item['report_text'])

            st.markdown("---")
            st.info(f"📁 히스토리 저장 경로: **{os.path.abspath(sub_folder)}**")

            c1, c2, c3 = st.columns(3)
            with c1:
                with open(doc_filename, "rb") as f:
                    st.download_button("📥 Word (.docx) 다운로드", f, file_name=os.path.basename(doc_filename), key="dl_docx_v89")
            with c2:
                with open(xlsx_filename, "rb") as f:
                    st.download_button("📥 Excel (.xlsx) 다운로드", f, file_name=os.path.basename(xlsx_filename), key="dl_xlsx_v89")
            with c3:
                with open(ppt_filename, "rb") as f:
                    st.download_button("📥 PowerPoint (.pptx) 다운로드", f, file_name=os.path.basename(ppt_filename), key="dl_pptx_v89")

        except Exception as e:
            st.error(f"오류가 발생했습니다: {str(e)}")
            st.exception(e)
