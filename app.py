"""
app.py
AI 포스터 생성기 — Flask 웹 서버
파일을 임시 보관 후 직접 다운로드 방식
"""

import os
import json
import base64
import uuid
import time
import threading
from io import BytesIO
from flask import (
    Flask, render_template, request, jsonify,
    send_file
)
from poster_engine import (
    SIZE_PRESETS, get_size_info, init_gemini,
    generate_image, generate_text_plan, create_poster_pptx,
)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "ai-poster-secret-key-change-me")

# ── 임시 메모리 저장소 ──
memory_store = {}

def get_store(sid):
    if sid not in memory_store:
        memory_store[sid] = {"created": time.time()}
    return memory_store[sid]

def clear_store(sid):
    if sid in memory_store:
        del memory_store[sid]

# ── 오래된 세션 자동 정리 (10분) ──
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
        sid = str(uuid.uuid4())

        if data.get("size_type") == "custom":
            size = get_size_info(
                custom_w=float(data["custom_width"]),
                custom_h=float(data["custom_height"]),
            )
        else:
            size = get_size_info(preset_key=data["preset"])

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
            "mood": data.get("mood", ""),
            "additional": data.get("additional", ""),
            "size_name": size["name"],
            "gemini_ratio": size["gemini_ratio"],
            "upload_descriptions": ", ".join(upload_descriptions) if upload_descriptions else "없음",
        }

        plan_text = generate_text_plan(client, poster_info)

        sketch_prompt = f"""Create a rough sketch/draft poster design.
Theme/Title: {data.get('title', '')}
Date: {data.get('date', '')}
Venue: {data.get('venue', '')}
Style/Mood: {data.get('mood', '')}
Content to include: {data.get('content', '')}
Additional creative requests: {data.get('additional', '')}
Uploaded materials: {', '.join(upload_descriptions) if upload_descriptions else 'none'}

IMPORTANT: This poster has an aspect ratio of {size['gemini_ratio']}.
Design accordingly — {"wide/horizontal layout" if size['width_cm'] > size['height_cm'] else "tall/vertical layout" if size['height_cm'] > size['width_cm'] else "square layout"}.

This is a DRAFT sketch. Include all text in Korean.
Show approximate layout with colors, text positions, and decorative elements.
Leave clear spaces where uploaded images (logos, photos) would be placed.
Make it look like a professional poster."""

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

        revise_prompt = f"""I have this poster design draft. Please revise it based on this feedback:

Feedback: {feedback}

Keep the overall theme for: {info['title']}
Style/Mood: {info['mood']}
Aspect ratio: {size['gemini_ratio']}
{"Wide/horizontal layout" if size['width_cm'] > size['height_cm'] else "Tall/vertical layout" if size['height_cm'] > size['width_cm'] else "Square layout"}

Apply the requested changes while maintaining the poster's quality.
All text in Korean."""

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
    """최종 생성 — 파일을 메모리에 보관하고 다운로드 URL 제공"""
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

        # 1) 완성본
        final_prompt = f"""Based on this draft, create the FINAL high-quality poster.
Keep the exact same layout, colors, style, and composition.
Make everything polished, crisp, and professional.
All Korean text must be accurate and readable.
Title: {info['title']}
Date: {info.get('date', '')}
Venue: {info.get('venue', '')}
Aspect ratio must be exactly {size['gemini_ratio']}."""

        reference_bytes = generate_image(
            client, prompt=final_prompt,
            aspect_ratio=size["gemini_ratio"],
            image_size="2K",
            ref_image_bytes=sketch,
        )

        # 2) 배경
        bg_prompt = """Remove ALL text from this poster image completely.
Keep everything else exactly the same — colors, illustrations, decorations, patterns, background.
Fill the text areas seamlessly with the surrounding background.
No text, no letters, no numbers, no words anywhere in the result."""

        background_bytes = generate_image(
            client, prompt=bg_prompt,
            aspect_ratio=size["gemini_ratio"],
            image_size="2K",
            ref_image_bytes=reference_bytes,
        )

        # 3) 텍스트 배치
        w = size["width_cm"]
        h = size["height_cm"]
        margin = min(w, h) * 0.05
        font_name = store.get("font_name", "맑은 고딕")

        texts = []
        if info.get("title"):
            texts.append({
                "content": info["title"],
                "x_cm": margin, "y_cm": h * 0.08,
                "width_cm": w - margin * 2, "height_cm": h * 0.15,
                "font_size_pt": max(20, min(72, int(min(w, h) * 1.2))),
                "font_name": font_name,
                "font_color": "FFFFFF", "bold": True, "align": "center",
            })
        if info.get("date"):
            texts.append({
                "content": info["date"],
                "x_cm": margin, "y_cm": h * 0.55,
                "width_cm": w - margin * 2, "height_cm": h * 0.06,
                "font_size_pt": max(12, min(36, int(min(w, h) * 0.6))),
                "font_name": font_name,
                "font_color": "FFFFFF", "bold": False, "align": "center",
            })
        if info.get("venue"):
            texts.append({
                "content": info["venue"],
                "x_cm": margin, "y_cm": h * 0.62,
                "width_cm": w - margin * 2, "height_cm": h * 0.06,
                "font_size_pt": max(12, min(36, int(min(w, h) * 0.6))),
                "font_name": font_name,
                "font_color": "FFFFFF", "bold": False, "align": "center",
            })
        if info.get("content"):
            texts.append({
                "content": info["content"],
                "x_cm": margin, "y_cm": h * 0.70,
                "width_cm": w - margin * 2, "height_cm": h * 0.20,
                "font_size_pt": max(10, min(24, int(min(w, h) * 0.4))),
                "font_name": font_name,
                "font_color": "FFFFFF", "bold": False, "align": "center",
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

        # 5) PPTX 생성
        pptx_buffer = create_poster_pptx(
            width_cm=w, height_cm=h,
            background_bytes=background_bytes,
            reference_bytes=reference_bytes,
            texts=texts,
            upload_images=upload_for_pptx,
        )

        # 다운로드용 파일을 메모리에 저장
        download_id = str(uuid.uuid4())[:8]
        store["downloads"] = {
            "pptx": pptx_buffer.read(),
            "reference": reference_bytes,
            "background": background_bytes,
        }
        store["download_id"] = download_id
        store["created"] = time.time()  # 타이머 리셋

        # 미리보기용 스케치만 base64로 (작은 사이즈)
        ref_b64 = base64.b64encode(reference_bytes).decode("utf-8")

        return jsonify({
            "success": True,
            "sid": sid,
            "download_id": download_id,
            "reference_b64": ref_b64,
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": str(e)}), 500


# ── 파일 다운로드 엔드포인트 ──

@app.route("/download/<sid>/<file_type>")
def download_file(sid, file_type):
    """메모리에서 직접 파일을 전송"""
    store = memory_store.get(sid)
    if not store or "downloads" not in store:
        return "파일이 만료되었습니다. 다시 생성해주세요.", 404

    downloads = store["downloads"]

    if file_type == "pptx":
        return send_file(
            BytesIO(downloads["pptx"]),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="poster.pptx",
        )
    elif file_type == "reference":
        return send_file(
            BytesIO(downloads["reference"]),
            mimetype="image/jpeg",
            as_attachment=True,
            download_name="reference.jpg",
        )
    elif file_type == "background":
        return send_file(
            BytesIO(downloads["background"]),
            mimetype="image/jpeg",
            as_attachment=True,
            download_name="background.jpg",
        )
    else:
        return "잘못된 요청입니다.", 400


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 50)
    print("  🎨 AI 포스터 생성기")
    print(f"  http://localhost:{port}")
    print("=" * 50)
    app.run(debug=True, host="0.0.0.0", port=port)
