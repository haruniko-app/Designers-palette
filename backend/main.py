import base64
import json
import functions_framework
import vertexai
from vertexai.preview.vision_models import ImageGenerationModel, Image as VertexImage
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter
from io import BytesIO
import math

# ==========================================
# 👇 あなたのプロジェクトIDに書き換えてください
PROJECT_ID = "slide-ai-tool"
# ==========================================
LOCATION = "us-central1"

vertexai.init(project=PROJECT_ID, location=LOCATION)

@functions_framework.http
def generate_image(request):
    if request.method == 'OPTIONS':
        headers = {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST',
            'Access-Control-Allow-Headers': 'Content-Type',
            'Access-Control-Max-Age': '3600'
        }
        return ('', 204, headers)

    headers = {'Access-Control-Allow-Origin': '*'}

    try:
        req = request.get_json(silent=True)
        if not req or 'image' not in req:
             return ({"error": "No image data"}, 400, headers)

        img_b64 = req['image'].split(",")[1] if "," in req['image'] else req['image']
        image_bytes = base64.b64decode(img_b64)
        original_pil = Image.open(BytesIO(image_bytes)).convert("RGB")
        
        mode = req.get('mode', 'extend')
        final_pil_image = None

        # =================================================================
        # モードA: 高画質化 (Smart Resize + Imagen 4.0)
        # =================================================================
        if mode == 'upscale':
            print("Mode: Upscale selected")
            factor_str = req.get('upscale_factor', '2x') # "2x" or "4x"
            factor_int = 4 if '4' in factor_str else 2
            
            # --- 1. サイズ制限チェックと自動調整 ---
            # Google Imagenの制限: 出力画素数が 17,000,000 (17MP) 以下であること
            LIMIT_PIXELS = 16500000 # 安全マージンをとって16.5MPに設定
            
            current_w, current_h = original_pil.size
            target_pixels = (current_w * factor_int) * (current_h * factor_int)
            
            if target_pixels > LIMIT_PIXELS:
                print(f"⚠️ Image too large for AI ({target_pixels} px). Resizing to fit limit...")
                
                # 縮小率を計算: sqrt(上限 / 現在の予定画素数)
                scale_ratio = math.sqrt(LIMIT_PIXELS / target_pixels)
                
                # 新しいサイズ (AIに入力するサイズ)
                new_input_w = int(current_w * scale_ratio)
                new_input_h = int(current_h * scale_ratio)
                
                # 高品質に縮小
                original_pil = original_pil.resize((new_input_w, new_input_h), Image.Resampling.LANCZOS)
                
                # バイトデータを更新
                buf = BytesIO()
                original_pil.save(buf, format="PNG")
                image_bytes = buf.getvalue()
                
                print(f"Resized input to: {new_input_w}x{new_input_h}")

            # --- 2. Google最新AI (Imagen 4.0) を実行 ---
            # Google仕様: 文字列 "x2" or "x4" が必要
            api_factor = "x2" if factor_int == 2 else "x4"
            
            try:
                print(f"Trying Imagen 4.0 Upscale with factor: {api_factor}")
                model = ImageGenerationModel.from_pretrained("imagen-4.0-upscale-preview")
                vertex_img = VertexImage(image_bytes)
                
                result_img = model.upscale_image(
                    image=vertex_img,
                    upscale_factor=api_factor
                )
                
                generated_result = result_img[0] if isinstance(result_img, list) else result_img
                
                if hasattr(generated_result, "_image_bytes"):
                    final_pil_image = Image.open(BytesIO(generated_result._image_bytes))
                    print("✅ Imagen 4.0 Upscale Success!")
                else:
                    raise ValueError("Invalid response format")

            except Exception as ai_error:
                # --- 3. それでもダメなら Pythonで救済 ---
                print(f"⚠️ AI Failed: {ai_error}. Switching to Python Lanczos.")
                orig_w, orig_h = original_pil.size
                new_w = orig_w * factor_int
                new_h = orig_h * factor_int
                
                resized_pil = original_pil.resize((new_w, new_h), Image.Resampling.LANCZOS)
                sharpener = ImageFilter.UnsharpMask(radius=1.5, percent=150, threshold=3)
                final_pil_image = resized_pil.filter(sharpener)
                enhancer = ImageEnhance.Contrast(final_pil_image)
                final_pil_image = enhancer.enhance(1.05)

        # =================================================================
        # モードB: AI拡張 (Imagen 2)
        # =================================================================
        else:
            # (前回のコードと同じなので省略なしで記載)
            print("Mode: Extend (Imagen 2) selected")
            expand_w_percent = int(req.get('expand_w', 0))
            expand_h_percent = int(req.get('expand_h', 0))
            prompt_text = req.get('prompt', "")
            
            orig_w, orig_h = original_pil.size
            add_w = int(orig_w * (expand_w_percent / 100) * 2)
            add_h = int(orig_h * (expand_h_percent / 100) * 2)
            new_w = orig_w + add_w
            new_h = orig_h + add_h
            
            if add_w == 0 and add_h == 0:
                 return ({"error": "拡張幅が0です"}, 400, headers)

            base_canvas = Image.new("RGB", (new_w, new_h), (255, 255, 255))
            paste_x = add_w // 2
            paste_y = add_h // 2
            base_canvas.paste(original_pil, (paste_x, paste_y))

            # マスク作成：小さめのマージンで境界をAIに任せる
            margin = 8  # AIが自然にブレンドできる最小限のマージン
            mask_canvas = Image.new("L", (new_w, new_h), 255)
            draw = ImageDraw.Draw(mask_canvas)
            draw.rectangle(
                (paste_x + margin, paste_y + margin, paste_x + orig_w - margin, paste_y + orig_h - margin),
                fill=0
            )

            buff_base = BytesIO()
            base_canvas.save(buff_base, format="PNG")
            vertex_base_img = VertexImage(buff_base.getvalue())

            buff_mask = BytesIO()
            mask_canvas.save(buff_mask, format="PNG")
            vertex_mask_img = VertexImage(buff_mask.getvalue())

            model = ImageGenerationModel.from_pretrained("imagegeneration@006")
            # Phase 2: プロンプト改善案A - 境界のシームレスなブレンドを強調
            full_prompt = "Perfect outpainting with seamless blending. Continue the image naturally with no visible edges, borders, or seams. Match existing style, colors, lighting, and details perfectly. Coherent extension. " + prompt_text

            images = model.edit_image(
                base_image=vertex_base_img,
                mask=vertex_mask_img,
                prompt=full_prompt,
                guidance_scale=50,
                number_of_images=1
            )
            result_bytes = images[0]._image_bytes
            final_pil_image = Image.open(BytesIO(result_bytes))

        # =================================================================
        # 結果返却
        # =================================================================
        if not final_pil_image:
             return ({"error": "Processing failed"}, 500, headers)

        out_buff = BytesIO()
        final_pil_image.save(out_buff, format="PNG")
        out_b64 = base64.b64encode(out_buff.getvalue()).decode('utf-8')

        return ({"result_image": "data:image/png;base64," + out_b64}, 200, headers)

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return ({"error": f"Server Error: {str(e)}"}, 500, headers)