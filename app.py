"""
app.py
AI 포스터 생성기 — Flask 웹 서버
v2: 텍스트 세분화, 스타일 프리셋, 3슬라이드, 재시도, 원본 다운로드
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

def clear_store(sid):
    if sid in memory_store:
        del memory_store[sid]

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
    """파일명에 사용 가능한 문자만 남기기"""
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

        # API 키 유효성 체크
        if not validate_api_key(client):
            return jsonify({"success": False, "error": "API 키가 유효하지 않습니다. 키를 다시 확인해주세요."}), 400

        sid = str(uuid.uuid4())

        if data.get("size_type") == "custom":
            size = get_size_info(
                custom_w=float(data["custom_width"]),
                custom_h=float(data["custom_height"]),
            )
        else:
            size = get_size_info(preset_key=data["preset"])

        # 스타일 처리
        style_key = data.get("style_preset", "custom")
        if style_key in STYLE_PRESETS and style_key != "custom":
            style_prompt_en = STYLE_PRESETS[style_key]["prompt_en"]
            style_name = STYLE_PRESETS[style_key]["name"]
        else:
            style_prompt_en = ""
            style_name = "커스텀"

        mood_text = data.get("mood", "")

        ref_bytes = None
        if "reference_image" in request.files:
            f = request.files["reference_image"]
            if f.filename:
                ref_bytes = f.read()

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

        font_name = "맑은 고딕"
        if "font_file" in request.files:
            f = request.files["font_file"]
            if f.filename:
                font_name = os.path.splitext(f.filename)[0]

        poster_info = {
            "title": data.get("title", ""),
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

        plan_text = generate_text_plan(client, poster_info)

        # 스타일 프롬프트 조합
        style_instruction = style_prompt_en
        if mood_text:
            style_instruction += f"\nAdditional mood/style: {mood_text}"

        orientation = "wide/horizontal layout" if size['width_cm'] > size['height_cm'] else "tall/vertical layout" if size['height_cm'] > size['width_cm'] else "square layout"

        sketch_prompt = f"""Create a poster design draft.
Theme/Title: {data.get('title', '')}
Date: {data.get('date', '')}
Venue: {data.get('venue', '')}
Organizer: {data.get('organizer', '')}
Content to include: {data.get('content', '')}
Additional creative requests: {data.get('additional', '')}
Uploaded materials: {', '.join(upload_descriptions) if upload_descriptions else 'none'}

DESIGN STYLE: {style_instruction}

IMPORTANT: This poster has an aspect ratio of {size['gemini_ratio']}.
Design accordingly — {orientation}.

Include all text in Korean.
Show approximate layout with colors, text positions, and decorative elements.
Leave clear spaces where uploaded images (logos, photos) would be placed.
Make it look like a professional poster.
For the bottom information area (contact, registration, organizer), keep the background clean so text can be overlaid later."""

        sketch_bytes = generate_image(
            client,
            prompt=sketch_prompt,
            aspect_ratio=size["gemini_ratio"],
            image_size="1K",
            ref_image_bytes=ref_bytes,
        )

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
            error_msg = "API 키가 유효하지 않습니다. 키를 다시 확인해주세요."
        elif "quota" in error_msg.lower():
            error_msg = "API 사용량 한도에 도달했습니다. 잠시 후 다시 시도해주세요."
        elif "safety" in error_msg.lower():
            error_msg = "AI 안전 필터에 의해 차단되었습니다. 내용을 수정해주세요."
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

        orientation = "Wide/horizontal layout" if size['width_cm'] > size['height_cm'] else "Tall/vertical layout" if size['height_cm'] > size['width_cm'] else "Square layout"

        revise_prompt = f"""I have this poster design draft. Please revise it based on this feedback:

Feedback: {feedback}

Keep the overall theme for: {info['title']}
DESIGN STYLE: {style_instruction}
Aspect ratio: {size['gemini_ratio']}
{orientation}

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
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/generate", methods=["POST"])
def api_generate():
    """최종 생성 — 3슬라이드 PPTX"""
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

        # 1) 완성본 (원본 — 텍스트 포함)
        final_prompt = f"""Based on this draft, create the FINAL high-quality poster.
Keep the exact same layout, colors, style, and composition.
Make everything polished, crisp, and professional.
All Korean text must be accurate and readable.
DESIGN STYLE: {style_instruction}
Title: {info['title']}
Date: {info.get('date', '')}
Venue: {info.get('venue', '')}
Organizer: {info.get('organizer', '')}
Aspect ratio must be exactly {size['gemini_ratio']}."""

        original_bytes = generate_image(
            client, prompt=final_prompt,
            aspect_ratio=size["gemini_ratio"],
            image_size="2K",
            ref_image_bytes=sketch,
        )

        # 2) 배경 (일러스트 라벨 유지, 안내 텍스트만 제거)
        bg_prompt = """Remove the informational text from this poster image.
Keep ALL illustration labels and decorative text that are part of the artwork (like category names on illustrations).
Remove ONLY the main title text, date, venue, contact information, registration details, and organizer text.
Fill those removed text areas seamlessly with the surrounding background.
Keep all illustrations, decorations, colors, and artistic elements exactly the same."""

        background_bytes = generate_image(
            client, prompt=bg_prompt,
            aspect_ratio=size["gemini_ratio"],
            image_size="2K",
            ref_image_bytes=original_bytes,
        )

        # 3) 텍스트 세분화 배치
        w = size["width_cm"]
        h = size["height_cm"]
        
        # PPTX 최대 크기 제한 적용
        MAX_CM = 142.0
        pptx_scale = 1.0
        if w > MAX_CM or h > MAX_CM:
            pptx_scale = min(MAX_CM / w, MAX_CM / h)
        
        # 텍스트 배치는 스케일 적용된 크기 기준
        pw = w * pptx_scale
        ph = h * pptx_scale
        margin = min(pw, ph) * 0.05
        font_name = store.get("font_name", "맑은 고딕")
        font_color = info.get("font_color", "FFFFFF")

        texts = []

        if info.get("title"):
            texts.append({
                "content": info["title"],
                "x_cm": margin, "y_cm": ph * 0.05,
                "width_cm": pw - margin * 2, "height_cm": ph * 0.12,
                "font_size_pt": max(24, min(72, int(min(pw, ph) * 1.2))),
                "font_name": font_name,
                "font_color": font_color, "bold": True, "align": "center",
            })

        if info.get("date"):
            texts.append({
                "content": info["date"],
                "x_cm": margin, "y_cm": ph * 0.18,
                "width_cm": pw - margin * 2, "height_cm": ph * 0.05,
                "font_size_pt": max(14, min(36, int(min(pw, ph) * 0.6))),
                "font_name": font_name,
                "font_color": font_color, "bold": True, "align": "center",
            })

        if info.get("venue"):
            texts.append({
                "content": info["venue"],
                "x_cm": margin, "y_cm": ph * 0.23,
                "width_cm": pw - margin * 2, "height_cm": ph * 0.05,
                "font_size_pt": max(12, min(30, int(min(pw, ph) * 0.5))),
                "font_name": font_name,
                "font_color": font_color, "bold": False, "align": "center",
            })

        if info.get("content"):
            content_lines = [line.strip() for line in info["content"].split('\n') if line.strip()]
            start_y = ph * 0.55
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

        if info.get("organizer"):
            texts.append({
                "content": f"주관: {info['organizer']}",
                "x_cm": margin, "y_cm": ph * 0.92,
                "width_cm": pw - margin * 2, "height_cm": ph * 0.05,
                "font_size_pt": max(12, min(28, int(min(pw, ph) * 0.5))),
                "font_name": font_name,
                "font_color": font_color, "bold": True, "align": "center",
            })

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
        # 고정 멘트 — 줄바꿈 기준으로 분리
        if info.get("content"):
            content_lines = [line.strip() for line in info["content"].split('\n') if line.strip()]
            start_y = h * 0.55
            line_height = h * 0.05
            for i, line in enumerate(content_lines):
                texts.append({
                    "content": line,
                    "x_cm": margin, "y_cm": start_y + (i * line_height),
                    "width_cm": w - margin * 2, "height_cm": line_height,
                    "font_size_pt": max(10, min(24, int(min(w, h) * 0.4))),
                    "font_name": font_name,
                    "font_color": font_color, "bold": False, "align": "center",
                })

        # 주관
        if info.get("organizer"):
            texts.append({
                "content": f"주관: {info['organizer']}",
                "x_cm": margin, "y_cm": h * 0.92,
                "width_cm": w - margin * 2, "height_cm": h * 0.05,
                "font_size_pt": max(12, min(28, int(min(w, h) * 0.5))),
                "font_name": font_name,
                "font_color": font_color, "bold": True, "align": "center",
            })

        # 4) 업로드 이미지 배치
        upload_for_pptx = []
        if store.get("upload_images"):
            for i, img in enumerate(store["upload_images"]):
                img_w = min(w * 0.15, 8)
                img_h = img_w
                upload_for_pptx.append({
                    "image_bytes": img["image_bytes"],
                    "description": img["description"],
                    "x_cm": w - margin - img_w - (i * (img_w + 1)),
                    "y_cm": h - margin - img_h,
                    "width_cm": img_w,
                    "height_cm": img_h,
                })

        # 5) PPTX 생성 (3슬라이드)
        poster_title = safe_filename(info.get("title", "poster"))
        pptx_buffer = create_poster_pptx(
            width_cm=w, height_cm=h,
            background_bytes=background_bytes,
            original_bytes=original_bytes,
            reference_bytes=original_bytes,
            texts=texts,
            upload_images=upload_for_pptx,
            title=poster_title,
        )

        # 다운로드용 파일을 메모리에 저장
        store["downloads"] = {
            "pptx": pptx_buffer.read(),
            "original": original_bytes,
            "background": background_bytes,
            "title": poster_title,
        }
        store["created"] = time.time()

        ref_b64 = base64.b64encode(original_bytes).decode("utf-8")

        return jsonify({
            "success": True,
            "sid": sid,
            "reference_b64": ref_b64,
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        error_msg = str(e)
        if "quota" in error_msg.lower():
            error_msg = "API 사용량 한도에 도달했습니다. 잠시 후 다시 시도해주세요."
        elif "safety" in error_msg.lower():
            error_msg = "AI 안전 필터에 의해 차단되었습니다."
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
    print("  🎨 AI 포스터 생성기 v2")
    print(f"  http://localhost:{port}")
    print("=" * 50)
    app.run(debug=True, host="0.0.0.0", port=port)
