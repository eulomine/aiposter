"""
poster_engine.py
포스터 생성 엔진 — 사이즈 관리, Gemini API 호출, PPTX 조립
모든 작업을 메모리에서 처리 (디스크 저장 없음)
"""

import base64
import math
from io import BytesIO
from google import genai
from google.genai import types
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image

# ============================================================
# 1. 사이즈 프리셋 & 비율 매칭
# ============================================================

GEMINI_RATIOS = [
    (1, 1), (3, 2), (2, 3), (3, 4), (4, 3),
    (4, 5), (5, 4), (9, 16), (16, 9), (21, 9),
    (9, 21), (1, 4), (4, 1), (1, 8), (8, 1),
]

SIZE_PRESETS = {
    "A3_landscape":   {"name": "A3 가로",       "width_cm": 42.0,   "height_cm": 29.7},
    "A3_portrait":    {"name": "A3 세로",       "width_cm": 29.7,   "height_cm": 42.0},
    "16:9_landscape": {"name": "16:9 가로",     "width_cm": 50.8,   "height_cm": 28.575},
    "16:9_portrait":  {"name": "16:9 세로",     "width_cm": 28.575, "height_cm": 50.8},
    "square":         {"name": "정방형 (1:1)",   "width_cm": 40.0,   "height_cm": 40.0},
    "banner_h":       {"name": "가로 배너",      "width_cm": 120.0,  "height_cm": 30.0},
    "banner_v":       {"name": "세로 배너",      "width_cm": 30.0,   "height_cm": 120.0},
    "hanging_banner": {"name": "가로 현수막",    "width_cm": 500.0,  "height_cm": 90.0},
}


def find_closest_gemini_ratio(width_cm: float, height_cm: float) -> str:
    """사용자 크기 → Gemini 지원 비율 중 가장 가까운 것"""
    target = width_cm / height_cm
    best_ratio = None
    best_diff = float('inf')
    for w, h in GEMINI_RATIOS:
        diff = abs(math.log(target) - math.log(w / h))
        if diff < best_diff:
            best_diff = diff
            best_ratio = (w, h)
    return f"{best_ratio[0]}:{best_ratio[1]}"


def get_size_info(preset_key=None, custom_w=None, custom_h=None):
    """프리셋 또는 커스텀 → {name, width_cm, height_cm, gemini_ratio}"""
    if preset_key and preset_key in SIZE_PRESETS:
        info = SIZE_PRESETS[preset_key].copy()
        info["gemini_ratio"] = find_closest_gemini_ratio(info["width_cm"], info["height_cm"])
        return info
    elif custom_w and custom_h:
        return {
            "name": f"커스텀 ({custom_w}×{custom_h}cm)",
            "width_cm": float(custom_w),
            "height_cm": float(custom_h),
            "gemini_ratio": find_closest_gemini_ratio(float(custom_w), float(custom_h)),
        }
    raise ValueError("사이즈를 지정해주세요.")


# ============================================================
# 2. Gemini API 호출
# ============================================================

def init_gemini(api_key: str):
    return genai.Client(api_key=api_key)


def generate_image(client, prompt, aspect_ratio, image_size="2K",
                   ref_image_bytes=None, model="gemini-3.1-flash-image"):
    """
    Gemini로 이미지를 생성하고 JPEG bytes를 반환.
    ref_image_bytes가 있으면 이미지를 참조하여 편집.
    """
    contents = []

    if ref_image_bytes:
        contents.append(
            types.Part.from_bytes(data=ref_image_bytes, mime_type="image/jpeg")
        )

    contents.append(types.Part.from_text(text=prompt))

    response = client.models.generate_content(
        model=model,
        contents=contents,
        config=types.GenerateContentConfig(
            response_modalities=["IMAGE", "TEXT"],
            response_mime_type="image/jpeg",
        ),
    )

    # 응답에서 이미지 추출
    for part in response.candidates[0].content.parts:
        if part.inline_data and part.inline_data.mime_type.startswith("image/"):
            return part.inline_data.data

    raise RuntimeError("Gemini가 이미지를 반환하지 않았습니다.")


def generate_text_plan(client, poster_info, model="gemini-2.5-flash"):
    """텍스트 모델로 디자인 계획 생성"""
    prompt = f"""당신은 전문 포스터 디자이너입니다. 다음 정보로 포스터 디자인을 계획해주세요.

포스터 정보:
- 제목: {poster_info.get('title', '')}
- 날짜: {poster_info.get('date', '')}
- 장소: {poster_info.get('venue', '')}
- 내용: {poster_info.get('content', '')}
- 분위기/스타일: {poster_info.get('mood', '')}
- 크기: {poster_info.get('size_name', '')} (비율: {poster_info.get('gemini_ratio', '')})
- AI 추가 표현 요청: {poster_info.get('additional', '')}
- 업로드 이미지: {poster_info.get('upload_descriptions', '없음')}

다음을 포함해서 답변해주세요:
1. 전체 분위기와 색감 설명
2. 배경 이미지 생성용 프롬프트 (영어, 상세하게)
3. 텍스트 배치 계획 (위치, 크기, 색상)
4. 업로드된 이미지(로고, 사진 등)의 권장 배치 위치와 크기
5. 추천 장식 요소"""

    response = client.models.generate_content(model=model, contents=prompt)
    return response.text


# ============================================================
# 3. PPTX 생성 (메모리에서 처리)
# ============================================================

def create_poster_pptx(
    width_cm, height_cm,
    background_bytes,
    reference_bytes,
    texts,
    upload_images=None,
):
    """
    PPTX를 메모리에서 생성하여 BytesIO로 반환.

    texts: [{"content", "x_cm", "y_cm", "width_cm", "height_cm",
             "font_size_pt", "font_name", "font_color", "bold", "align"}, ...]

    upload_images: [{"image_bytes", "description", "x_cm", "y_cm",
                     "width_cm", "height_cm"}, ...]
    """
    prs = Presentation()
    prs.slide_width = Cm(width_cm)
    prs.slide_height = Cm(height_cm)

    blank_layout = prs.slide_layouts[6]

    # ── 슬라이드 1: 편집용 포스터 ──
    slide1 = prs.slides.add_slide(blank_layout)

    # 배경 이미지
    slide1.shapes.add_picture(
        BytesIO(background_bytes), Cm(0), Cm(0),
        width=Cm(width_cm), height=Cm(height_cm)
    )

    # 텍스트 박스들
    for t in texts:
        txBox = slide1.shapes.add_textbox(
            Cm(t["x_cm"]), Cm(t["y_cm"]),
            Cm(t["width_cm"]), Cm(t["height_cm"]),
        )
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = t["content"]
        p.font.size = Pt(t.get("font_size_pt", 24))
        p.font.name = t.get("font_name", "맑은 고딕")
        p.font.bold = t.get("bold", False)
        color_hex = t.get("font_color", "FFFFFF")
        p.font.color.rgb = RGBColor.from_string(color_hex)
        align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
        p.alignment = align_map.get(t.get("align", "center"), PP_ALIGN.CENTER)

    # 업로드 이미지들 (로고, 강사사진 등)
    if upload_images:
        for img in upload_images:
            slide1.shapes.add_picture(
                BytesIO(img["image_bytes"]),
                Cm(img["x_cm"]), Cm(img["y_cm"]),
                width=Cm(img["width_cm"]), height=Cm(img["height_cm"]),
            )

    # ── 슬라이드 2: 참고용 완성 이미지 ──
    slide2 = prs.slides.add_slide(blank_layout)
    slide2.shapes.add_picture(
        BytesIO(reference_bytes), Cm(0), Cm(0),
        width=Cm(width_cm), height=Cm(height_cm)
    )
    note_w = min(width_cm - 2, 25)
    note_box = slide2.shapes.add_textbox(Cm(1), Cm(1), Cm(note_w), Cm(3))
    note_p = note_box.text_frame.paragraphs[0]
    note_p.text = "📌 참고용 이미지입니다. 편집 완료 후 이 슬라이드를 삭제하세요."
    note_p.font.size = Pt(14)
    note_p.font.color.rgb = RGBColor(255, 0, 0)

    # BytesIO로 반환
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output
