#!/usr/bin/env python3
"""
HKTVmall 商戶招商 PPT 生成器 v3
基於 Short Version 模板 (22頁)

使用方法:
  python hktv_ppt_generator.py                                    # 生成示例 PPT
  python hktv_ppt_generator.py --serve                           # 啟動 Flask API 服務
  python hktv_ppt_generator.py --merchant "ABC電器" --pains "物流成本高,平台費用不透明"

指令 1: 頁面 [1, 3-22] 保留原有內容及圖示，頁 1 加入商戶名稱
指令 2: 頁面 [2] 根據痛點選擇動態生成

API 端點: POST /generate
  JSON body: {
    "merchant_name": "ABC電器有限公司",
    "category": "本地" | "海外" | "服務",
    "pain_points": ["物流成本高", "平台費用不透明"],
    "contact_email": "aog.merc@hktv.com.hk"
  }
"""

import io
import os
import sys
import json
import argparse
from pathlib import Path

from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor as RgbColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

# 預設模板路徑
DEFAULT_TEMPLATE = "template.pptx"

# ============================================================
# 模板數據
# ============================================================

# 痛點與解決方案
PAIN_SOLUTIONS = {
    "物流成本高": {
        "pain": "物流成本高",
        "solution": "全港最大住宅配送網絡，350+輛貨車，500+取貨點，大幅降低物流成本"
    },
    "平台費用不透明": {
        "pain": "平台費用不透明",
        "solution": "佣金按分類碼固定比例計算，每月15個工作日結算，無隱藏收費"
    },
    "缺乏電商經驗": {
        "pain": "缺乏電商經驗",
        "solution": "電商學院提供完整培訓，MMS後台操作、產品上架、廣告投放一對一支援"
    },
    "擔心庫存風險": {
        "pain": "擔心庫存風險",
        "solution": "GREEN LAB臨期百貨計劃，3PL靈活存貨管理，代售模式先售後買降低風險"
    },
    "希望提升品牌知名度": {
        "pain": "希望提升品牌知名度",
        "solution": "每年過億廣告投放，TVB黃金時段，58個MTR站，23列全車身，260萬+買家"
    },
    "希望拓展年輕客群": {
        "pain": "希望拓展年輕客群",
        "solution": "精準CRM定位，25-40歲核心消費群，每季度4.7x購買頻率的忠實客戶"
    },
    "希望增加線上曝光": {
        "pain": "希望增加線上曝光",
        "solution": "關鍵字廣告、橫幅廣告、首頁推廣位，SEO及站內搜尋排名優化"
    }
}

# ============================================================
# PPT 生成函數
# ============================================================

def add_text_to_slide(slide, text, x, y, w, h, size, bold=False, color=None, align=PP_ALIGN.LEFT):
    """添加文字框"""
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    para = tf.paragraphs[0]
    para.text = text
    para.font.size = Pt(size)
    para.font.bold = bold
    if color:
        para.font.color.rgb = RgbColor(*color)
    para.alignment = align
    return shape

def find_slide_with_text(slide, search_text):
    """在幻燈片中查找包含指定文字的形狀"""
    for shape in slide.shapes:
        if shape.has_text_frame:
            if search_text in shape.text_frame.text:
                return shape
    return None

# ============================================================
# 主 PPT 生成函數
# ============================================================

def generate_merchant_ppt(merchant_name=None, category="本地", pain_points=None, contact_email="aog.merc@hktv.com.hk", output_path=None, template_path=None):
    """
    生成商戶招商 PPT - 保留模板原有圖示

    指令 1: 頁面 [1, 3-22] 保留原有內容及圖示，頁 1 加入商戶名稱
    指令 2: 頁面 [2] 根據痛點選擇動態生成
    """
    if pain_points is None:
        pain_points = []

    # 確定模板路徑
    if template_path is None:
        # 嘗試多個可能的位置
        possible_paths = [
            DEFAULT_TEMPLATE,
            os.path.join(os.path.dirname(__file__), "template.pptx"),
            os.path.join(os.path.dirname(__file__), "2026 HKTVmall New Merchant Acquisition Deck_Short Verson [CHI].pptx"),
        ]
        for p in possible_paths:
            if os.path.exists(p):
                template_path = p
                break
        if template_path is None:
            raise FileNotFoundError(f"找不到模板檔案，請確認 template.pptx 在同一目錄")

    # 開啟模板
    prs = Presentation(template_path)

    # ===== 第1頁: 封面 - 加入商戶名稱 =====
    if merchant_name and len(prs.slides) >= 1:
        slide1 = prs.slides[0]
        # 在現有內容下方添加商戶名稱
        add_text_to_slide(
            slide1,
            f"「{merchant_name}」專屬方案",
            0.5, 4.5, 9, 0.6, 24, bold=True, color=(230, 0, 18), align=PP_ALIGN.CENTER
        )

    # ===== 第2頁: 痛點與解決方案 - 根據選擇動態生成 =====
    if pain_points and len(prs.slides) >= 2:
        slide2 = prs.slides[1]
        # 標題 "商戶資料【主要業務】和【痛點與解決方案】"
        # 查找並修改標題，或在頂部添加新文字框
        title = find_slide_with_text(slide2, "商戶資料")
        if title:
            # 更新現有標題
            title.text_frame.paragraphs[0].text = "商戶資料"
        else:
            # 添加新標題
            add_text_to_slide(slide2, "商戶資料", 0.5, 0.3, 9, 0.5, 26, bold=True, color=(34, 34, 34), align=PP_ALIGN.CENTER)

        # 添加痛點與解決方案
        y = 1.0
        for pain_key in pain_points:
            if pain_key in PAIN_SOLUTIONS:
                data = PAIN_SOLUTIONS[pain_key]
                # 痛點
                add_text_to_slide(slide2, f"❌ {data['pain']}", 0.8, y, 4, 0.35, 12, bold=True, color=(204, 0, 0))
                # 解決方案
                add_text_to_slide(slide2, f"✓ {data['solution']}", 4.8, y, 4.7, 0.35, 10, color=(0, 128, 0))
                y += 0.55

    # 保存
    if output_path:
        prs.save(output_path)
        print(f"PPT 已保存至: {output_path}")
    else:
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        return output

# ============================================================
# Flask API
# ============================================================

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate():
    """PPT 生成 API"""
    try:
        data = request.get_json()
        merchant_name = data.get('merchant_name', '')
        category = data.get('category', '本地')
        pain_points = data.get('pain_points', [])
        contact_email = data.get('contact_email', 'aog.merc@hktv.com.hk')

        ppt_file = generate_merchant_ppt(
            merchant_name=merchant_name,
            category=category,
            pain_points=pain_points,
            contact_email=contact_email
        )

        return send_file(
            ppt_file,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=f'HKTVmall_招商方案_{merchant_name}.pptx'
        )
    except FileNotFoundError as e:
        return jsonify({'error': str(e)}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

# ============================================================
# 主程序
# ============================================================

def main():
    parser = argparse.ArgumentParser(description='HKTVmall 商戶招商 PPT 生成器 v3 (保留模板圖示)')
    parser.add_argument('--serve', action='store_true', help='啟動 Flask API 服務')
    parser.add_argument('--output', '-o', help='輸出文件路徑')
    parser.add_argument('--template', '-t', help='模板路徑')
    parser.add_argument('--merchant', '-m', help='商戶名稱')
    parser.add_argument('--category', '-c', choices=['本地', '海外', '服務'], default='本地', help='招商計劃')
    parser.add_argument('--pains', '-p', help='痛點（逗號分隔）')
    parser.add_argument('--email', '-e', default='aog.merc@hktv.com.hk', help='聯絡電郵')
    parser.add_argument('--host', default='0.0.0.0', help='API 主機')
    parser.add_argument('--port', type=int, default=5000, help='API 端口')

    args = parser.parse_args()

    if args.serve:
        print(f"啟動 API 服務: http://{args.host}:{args.port}")
        app.run(host=args.host, port=args.port, debug=True)
    else:
        pain_points = []
        if args.pains:
            pain_points = [p.strip() for p in args.pains.split(',')]

        output_path = args.output or f'HKTVmall_招商方案_{args.merchant or "示例"}.pptx'
        print(f"生成 PPT: {output_path}")
        print(f"  商戶: {args.merchant or '示例'}")
        print(f"  痛點: {pain_points or '無'}")
        print(f"  模板: {args.template or 'template.pptx'}")

        try:
            generate_merchant_ppt(
                merchant_name=args.merchant,
                category=args.category,
                pain_points=pain_points,
                contact_email=args.email,
                output_path=output_path,
                template_path=args.template
            )
        except FileNotFoundError as e:
            print(f"錯誤: {e}")
            print("請確保 template.pptx 在同一目錄，或使用 --template 指定路徑")

if __name__ == '__main__':
    main()
