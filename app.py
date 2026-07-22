"""
app.py
AI 포스터 생성기 — Flask 웹 서버
v4: 절대 규칙, 부제 지원, 배경 프롬프트 개선, blank_layout 수정,
    텍스트 중복 제거, PPTX 크기 제한, 인쇄 크기 안내
"""

import os
import json
import base64
import uuid
import time
import threading
import re
from io import BytesIO
from flask import (
    Flask, render_template, request, jsonify,
    send_file
)
from poster_engine import (
    SIZE_PRESETS, STYLE_PRESETS, get_size_info, init_gemini,
    validate_api_key, generate_image, generate_text_plan,
    create_poster_pptx,
)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "ai-poster-secret-key-change-me")

memory_store = {}


def get_store(sid):
    if sid not in memory_store:
        memory_store[sid] = {"created": time.time()}
    return memory_store[sid]


def cleanup_old_sessions():
    while True:
        time.sleep(60)
        now = time.time()
        expired = [sid for sid, data in memory_store.items()
                   if now - data.get("created", now) > 600]
        for sid in expired:
            del memory_store[sid]


cleanup_thread = threading.Thread(target=cleanup_old_sessions, daemon=True)
cleanup_thread.start()


def safe_filename(text):
    cleaned = re.sub(r'[\\/*?:"<>|]', '', text)
    return cleaned[:50].strip() or "poster"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/plan", methods=["POST"])
def api_plan():
    try:
        data = request.form.to_dict()

        api_key = data.get("api_key", "").strip()
        if not api_key:
            return jsonify({"success": False, "error": "Gemini API 키를 입력해주세요."}), 400

        client = init_gemini(api_key)

        if not validate_api_key(client):
            return jsonify({"success": False, "error": "API 키가 유효하지 않습니다."}), 400

        sid = str(uuid.uuid4())

        # 사이즈
        if data.get("size_type") == "custom":
            size = get_size_info(
                custom_w=float(data["custom_width"]),
                custom_h=float(data["custom_height"]),
            )
        else:
            size = get_size_info(preset_key=data["preset"])

        # 스타일
        style_key = data.get("style_preset", "custom")
        if style_key in STYLE_PRESETS and style_key != "custom":
            style_prompt_en = STYLE_PRESETS[style_key]["prompt_en"]
            style_name = STYLE_PRESETS[style_key]["name"]
        else:
            style_prompt_en = ""
            style_name = "커스텀"

        mood_text = data.get("mood", "")

        # 참고 이미지
        ref_bytes = None
        if "reference_image" in request.files:
            f = request.files["reference_image"]
            if f.filename:
                ref_bytes = f.read()

        # 업로드 이미지
        upload_images = []
        upload_descriptions = []
        idx = 0
        while True:
            file_key = f"upload_image_{idx}"
            desc_key = f"upload_desc_{idx}"
            if file_key not in request.files:
                break
            f = request.files[file_key]
            if f.filename:
                img_bytes = f.read()
                desc = data.get(desc_key, f"이미지 {idx+1}")
                upload_images.append({
                    "image_bytes": img_bytes,
                    "description": desc,
                    "filename": f.filename,
                })
                upload_descriptions.append(f"{desc} ({f.filename})")
            idx += 1

        # 폰트
        font_name = "맑은 고딕"
        if "font_file" in request.files:
            f = request.files["font_file"]
            if f.filename:
                font_name = os.path.splitext(f.filename)[0]

        # poster_info 구성 — 부제 포함
        poster_info = {
            "title": data.get("title", ""),
            "subtitle": data.get("subtitle", ""),
            "date": data.get("date", ""),
            "venue": data.get("venue", ""),
            "content": data.get("content", ""),
            "organizer": data.get("organizer", ""),
            "mood": mood_text,
            "style_prompt_en": style_prompt_en,
            "style_name": style_name,
            "additional": data.get("additional", ""),
            "size_name": size["name"],
            "gemini_ratio": size["gemini_ratio"],
            "upload_descriptions": ", ".join(upload_descriptions) if upload_descriptions else "없음",
            "font_color": data.get("font_color", "FFFFFF"),
        }

        # 텍스트 계획
        plan_text = generate_text_plan(client, poster_info)

        # ── 스케치 프롬프트 ──
        style_instruction = style_prompt_en
        if mood_text:
            style_instruction += f"\nAdditional mood/style: {mood_text}"

        orientation = (
            "wide/horizontal layout" if size['width_cm'] > size['height_cm']
            else "tall/vertical layout" if size['height_cm'] > size['width_cm']
            else "square layout"
        )

        # 입력된 항목만 프롬프트에 포함
        has_uploads = len(upload_images) > 0
        if has_uploads:
            upload_prompt_section = f"Uploaded materials: {', '.join(upload_descriptions)}\nLeave appropriate spaces where these uploaded images would be placed."
        else:
            upload_prompt_section = "No uploaded images. Do NOT leave empty spaces for logos, photos, or QR codes. Fill the entire poster with graphics and design elements."

        info_lines = []
        if data.get('title'):
            info_lines.append(f"Title: {data['title']}")
        if data.get('subtitle'):
            info_lines.append(f"Subtitle: {data['subtitle']}")
        if data.get('date'):
            info_lines.append(f"Date: {data['date']}")
        if data.get('venue'):
            info_lines.append(f"Venue: {data['venue']}")
        if data.get('organizer'):
            info_lines.append(f"Organizer: {data['organizer']}")
        if data.get('content'):
            info_lines.append(f"Content to include: {data['content']}")
        if data.get('additional'):
            info_lines.append(f"Additional creative requests: {data['additional']}")
        info_block = "\n".join(info_lines)

        sketch_prompt = f"""Create a poster design draft.
{info_block}
{upload_prompt_section}

ABSOLUTE RULE: Do NOT invent or add any information that is not provided above.
Do NOT create fake phone numbers, email addresses, websites, registration methods, or contact details.
Only include information that is explicitly given.

DESIGN STYLE: {style_instruction}

IMPORTANT: This poster has an aspect ratio of {size['gemini_ratio']}.
Design accordingly — {orientation}.

Include all text in Korean.
Show approximate layout with colors, text positions, and decorative elements.
Make it look like a professional poster.
For the bottom information area, keep the background clean so text can be overlaid later."""

        sketch_bytes = generate_image(
            client,
            prompt=sketch_prompt,
            aspect_ratio=size["gemini_ratio"],
            image_size="1K",
            ref_image_bytes=ref_bytes,
        )

        # 세션 저장
        store = get_store(sid)
        store["api_key"] = api_key
        store["poster_info"] = poster_info
        store["size"] = size
        store["plan_text"] = plan_text
        store["sketch_bytes"] = sketch_bytes
        store["upload_images"] = upload_images
        store["font_name"] = font_name
        store["ref_bytes"] = ref_bytes

        sketch_b64 = base64.b64encode(sketch_bytes).decode("utf-8")

        return jsonify({
            "success": True,
            "sid": sid,
            "plan_text": plan_text,
            "sketch_b64": sketch_b64,
            "size": {
                "name": size["name"],
                "width_cm": size["width_cm"],
                "height_cm": size["height_cm"],
                "gemini_ratio": size["gemini_ratio"],
            },
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        error_msg = str(e)
        if "API key not valid" in error_msg:
            error_msg = "API 키가 유효하지 않습니다."
        elif "quota" in error_msg.lower():
            error_msg = "API 사용량 한도에 도달했습니다. 잠시 후 다시 시도해주세요."
        elif "safety" in error_msg.lower():
            error_msg = "AI 안전 필터에 의해 차단되었습니다."
        elif "503" in error_msg or "overloaded" in error_msg.lower():
            error_msg = "Gemini 서버가 일시적으로 과부하 상태입니다. 1~2분 후 다시 시도해주세요."
        return jsonify({"success": False, "error": error_msg}), 500


@app.route("/api/revise", methods=["POST"])
def api_revise():
    try:
        data = request.form.to_dict()
        sid = data["sid"]

        store = get_store(sid)
        if not store or "api_key" not in store:
            return jsonify({"success": False, "error": "세션이 만료되었습니다. 처음부터 다시 시작해주세요."}), 400

        client = init_gemini(store["api_key"])
        size = store["size"]
        info = store["poster_info"]
        feedback = data["feedback"]

        style_instruction = info.get("style_prompt_en", "")
        if info.get("mood"):
            style_instruction += f"\nMood: {info['mood']}"

        orientation = (
            "Wide/horizontal layout" if size['width_cm'] > size['height_cm']
            else "Tall/vertical layout" if size['height_cm'] > size['width_cm']
            else "Square layout"
        )

        revise_prompt = f"""I have this poster design draft. Please revise it based on this feedback:

Feedback: {feedback}

Keep the overall theme for: {info.get('title', '')}
Subtitle: {info.get('subtitle', '')}
DESIGN STYLE: {style_instruction}
Aspect ratio: {size['gemini_ratio']}
{orientation}

ABSOLUTE RULE: Do NOT invent or add any information that was not in the original design.
Do NOT create fake phone numbers, emails, websites, or contact details.

Apply the requested changes while maintaining the poster's quality.
All text in Korean.
Keep the bottom area clean for text overlay."""

        new_sketch = generate_image(
            client,
            prompt=revise_prompt,
            aspect_ratio=size["gemini_ratio"],
            image_size="1K",
            ref_image_bytes=store["sketch_bytes"],
        )

        store["sketch_bytes"] = new_sketch
        sketch_b64 = base64.b64encode(new_sketch).decode("utf-8")

        return jsonify({
            "success": True,
            "sid": sid,
            "sketch_b64": sketch_b64,
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        error_msg = str(e)
        if "503" in error_msg or "overloaded" in error_msg.lower():
            error_msg = "Gemini 서버가 일시적으로 과부하 상태입니다. 1~2분 후 다시 시도해주세요."
        return jsonify({"success": False, "error": error_msg}), 500


@app.route("/api/generate", methods=["POST"])
def api_generate():
    try:
        data = request.form.to_dict()
        sid = data["sid"]

        store = get_store(sid)
        if not store or "api_key" not in store:
            return jsonify({"success": False, "error": "세션이 만료되었습니다."}), 400

        client = init_gemini(store["api_key"])
        size = store["size"]
        info = store["poster_info"]
        sketch = store["sketch_bytes"]

        style_instruction = info.get("style_prompt_en", "")
        if info.get("mood"):
            style_instruction += f"\nMood: {info['mood']}"

        # ── 1) 완성본 (원본 — 텍스트 포함) ──
        final_info_lines = []
        if info.get('title'):
            final_info_lines.append(f"Title: {info['title']}")
        if info.get('subtitle'):
            final_info_lines.append(f"Subtitle: {info['subtitle']}")
        if info.get('date'):
            final_info_lines.append(f"Date: {info['date']}")
        if info.get('venue'):
            final_info_lines.append(f"Venue: {info['venue']}")
        if info.get('organizer'):
            final_info_lines.append(f"Organizer: {info['organizer']}")
        if info.get('content'):
            final_info_lines.append(f"Content: {info['content']}")
        final_info_block = "\n".join(final_info_lines)

        final_prompt = f"""Based on this draft, create the FINAL high-quality poster.
Keep the exact same layout, colors, style, and composition.
Make everything polished, crisp, and professional.
All Korean text must be accurate and readable.

ABSOLUTE RULE: Only include the following information. Do NOT add phone numbers, emails, websites, or any details not listed below.
{final_info_block}

DESIGN STYLE: {style_instruction}
Aspect ratio must be exactly {size['gemini_ratio']}."""

        original_bytes = generate_image(
            client, prompt=final_prompt,
            aspect_ratio=size["gemini_ratio"],
            image_size="2K",
            ref_image_bytes=sketch,
        )

        # ── 2) 배경 (모든 텍스트 제거, 디자인 요소 보존) ──
        bg_prompt = """Remove ALL text from this poster image.
This includes: titles, subtitles, dates, venue names, organizer names, category labels, and any Korean/English text.
IMPORTANT RULES:
- Remove ONLY the text characters themselves.
- Keep all background design elements EXACTLY as they are: boxes, frames, panels, shapes, colors, textures, illustrations, decorative patterns.
- If text was inside a colored box or panel, remove the text but keep the box/panel with its original color and shape intact.
- Fill the areas where text was removed with the SAME background that was behind the text.
- Do NOT change any colors, do NOT replace design elements, do NOT add new elements.
- The result should look like the same poster design but completely blank — ready for text to be overlaid."""

        background_bytes = generate_image(
            client, prompt=bg_prompt,
            aspect_ratio=size["gemini_ratio"],
            image_size="2K",
            ref_image_bytes=original_bytes,
        )

        # ── 3) 텍스트 배치 (스케일 적용) ──
        w = size["width_cm"]
        h = size["height_cm"]

        MAX_CM = 142.0
        pptx_scale = 1.0
        if w > MAX_CM or h > MAX_CM:
            pptx_scale = min(MAX_CM / w, MAX_CM / h)

        pw = w * pptx_scale
        ph = h * pptx_scale
        margin = min(pw, ph) * 0.05
        font_name = store.get("font_name", "맑은 고딕")
        font_color = info.get("font_color", "FFFFFF")

        texts = []

        # 현재 Y 위치 추적 — 입력된 항목만 배치
        current_y = ph * 0.05

        # 제목
        if info.get("title"):
            title_h = ph * 0.10
            texts.append({
                "content": info["title"],
                "x_cm": margin, "y_cm": current_y,
                "width_cm": pw - margin * 2, "height_cm": title_h,
                "font_size_pt": max(24, min(72, int(min(pw, ph) * 1.2))),
                "font_name": font_name,
                "font_color": font_color, "bold": True, "align": "center",
            })
            current_y += title_h

        # 부제
        if info.get("subtitle"):
            sub_h = ph * 0.08
            texts.append({
                "content": info["subtitle"],
                "x_cm": margin, "y_cm": current_y,
                "width_cm": pw - margin * 2, "height_cm": sub_h,
                "font_size_pt": max(18, min(54, int(min(pw, ph) * 0.9))),
                "font_name": font_name,
                "font_color": font_color, "bold": True, "align": "center",
            })
            current_y += sub_h

        # 날짜
        if info.get("date"):
            date_h = ph * 0.05
            current_y += ph * 0.02  # 약간 간격
            texts.append({
                "content": info["date"],
                "x_cm": margin, "y_cm": current_y,
                "width_cm": pw - margin * 2, "height_cm": date_h,
                "font_size_pt": max(14, min(36, int(min(pw, ph) * 0.6))),
                "font_name": font_name,
                "font_color": font_color, "bold": True, "align": "center",
            })
            current_y += date_h

        # 장소
        if info.get("venue"):
            venue_h = ph * 0.05
            texts.append({
                "content": info["venue"],
                "x_cm": margin, "y_cm": current_y,
                "width_cm": pw - margin * 2, "height_cm": venue_h,
                "font_size_pt": max(12, min(30, int(min(pw, ph) * 0.5))),
                "font_name": font_name,
                "font_color": font_color, "bold": False, "align": "center",
            })
            current_y += venue_h

        # 고정 멘트 — 줄바꿈 기준으로 각각 별도 텍스트 박스
        if info.get("content"):
            content_lines = [line.strip() for line in info["content"].split('\n') if line.strip()]
            start_y = max(current_y + ph * 0.05, ph * 0.50)
            line_height = ph * 0.05
            for i, line in enumerate(content_lines):
                texts.append({
                    "content": line,
                    "x_cm": margin, "y_cm": start_y + (i * line_height),
                    "width_cm": pw - margin * 2, "height_cm": line_height,
                    "font_size_pt": max(10, min(24, int(min(pw, ph) * 0.4))),
                    "font_name": font_name,
                    "font_color": font_color, "bold": False, "align": "center",
                })

        # 주관
        if info.get("organizer"):
            texts.append({
                "content": f"주관: {info['organizer']}",
                "x_cm": margin, "y_cm": ph * 0.92,
                "width_cm": pw - margin * 2, "height_cm": ph * 0.05,
                "font_size_pt": max(12, min(28, int(min(pw, ph) * 0.5))),
                "font_name": font_name,
                "font_color": font_color, "bold": True, "align": "center",
            })

        # ── 4) 업로드 이미지 배치 (있을 때만) ──
        upload_for_pptx = []
        if store.get("upload_images"):
            for i, img in enumerate(store["upload_images"]):
                img_w = min(pw * 0.15, 8)
                img_h = img_w
                upload_for_pptx.append({
                    "image_bytes": img["image_bytes"],
                    "description": img["description"],
                    "x_cm": pw - margin - img_w - (i * (img_w + 1)),
                    "y_cm": ph - margin - img_h,
                    "width_cm": img_w,
                    "height_cm": img_h,
                })

        # ── 5) PPTX 생성 ──
        poster_title = safe_filename(info.get("title", "poster"))

        actual_w = size["width_cm"] if pptx_scale < 1.0 else None
        actual_h = size["height_cm"] if pptx_scale < 1.0 else None

        pptx_buffer = create_poster_pptx(
            width_cm=w, height_cm=h,
            background_bytes=background_bytes,
            original_bytes=original_bytes,
            reference_bytes=original_bytes,
            texts=texts,
            upload_images=upload_for_pptx,
            title=poster_title,
            actual_width_cm=actual_w,
            actual_height_cm=actual_h,
        )

        store["downloads"] = {
            "pptx": pptx_buffer.read(),
            "original": original_bytes,
            "background": background_bytes,
            "title": poster_title,
        }
        store["created"] = time.time()

        ref_b64 = base64.b64encode(original_bytes).decode("utf-8")

        print_info = None
        if pptx_scale < 1.0:
            print_info = {
                "actual_width_cm": size["width_cm"],
                "actual_height_cm": size["height_cm"],
                "pptx_width_cm": round(w * pptx_scale, 2),
                "pptx_height_cm": round(h * pptx_scale, 2),
                "scale_percent": round(pptx_scale * 100, 1),
            }

        return jsonify({
            "success": True,
            "sid": sid,
            "reference_b64": ref_b64,
            "print_info": print_info,
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        error_msg = str(e)
        if "quota" in error_msg.lower():
            error_msg = "API 사용량 한도에 도달했습니다."
        elif "safety" in error_msg.lower():
            error_msg = "AI 안전 필터에 의해 차단되었습니다."
        elif "503" in error_msg or "overloaded" in error_msg.lower():
            error_msg = "Gemini 서버 과부하입니다. 1~2분 후 다시 시도해주세요."
        return jsonify({"success": False, "error": error_msg}), 500


@app.route("/download/<sid>/<file_type>")
def download_file(sid, file_type):
    store = memory_store.get(sid)
    if not store or "downloads" not in store:
        return "파일이 만료되었습니다. 다시 생성해주세요.", 404

    downloads = store["downloads"]
    title = downloads.get("title", "poster")

    if file_type == "pptx":
        return send_file(
            BytesIO(downloads["pptx"]),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=f"{title}.pptx",
        )
    elif file_type == "original":
        return send_file(
            BytesIO(downloads["original"]),
            mimetype="image/jpeg",
            as_attachment=True,
            download_name=f"{title}_원본.jpg",
        )
    elif file_type == "background":
        return send_file(
            BytesIO(downloads["background"]),
            mimetype="image/jpeg",
            as_attachment=True,
            download_name=f"{title}_배경.jpg",
        )
    else:
        return "잘못된 요청입니다.", 400


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 50)
    print("  🎨 AI 포스터 생성기 v4")
    print(f"  http://localhost:{port}")
    print("=" * 50)
    app.run(debug=True, host="0.0.0.0", port=port)
