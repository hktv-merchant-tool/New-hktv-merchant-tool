#!/usr/bin/env python3
"""
HKTVmall 商戶招商 PPT 生成器 v3
基於 Short Version 模板 (22頁)

使用方式:
  python hktv_ppt_generator.py --merchant "ABC電器" --pains "物流成本高,平台費用不透明"

指令 1: 頁面 [1, 3-22] 保留原有內容及圖示，頁 1 加入商戶名稱
指令 2: 頁面 [2] 根據痛點選擇動態生成

模板路徑: C:\\Users\\cylee\\Desktop\\2026 HKTVmall New Merchant Acquisition Deck_Short Verson [CHI].pptx
"""

import io
import os
import sys
import argparse

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor as RgbColor
from pptx.enum.text import PP_ALIGN

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

# 預設模板路徑
DEFAULT_TEMPLATE = "/mnt/c/Users/cylee/Desktop/2026 HKTVmall New Merchant Acquisition Deck_Short Verson [CHI].pptx"

def generate_merchant_ppt(merchant_name=None, pain_points=None, output_path=None, template_path=None):
    """
    生成商戶招商 PPT - 保留模板原有圖示

    指令 1: 頁面 [1, 3-22] 保留原有內容及圖示，頁 1 加入商戶名稱
    指令 2: 頁面 [2] 根據痛點選擇動態生成
    """
    if pain_points is None:
        pain_points = []

    # 確定模板路徑
    if template_path is None:
        template_path = DEFAULT_TEMPLATE

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"找不到模板檔案: {template_path}")

    # 開啟模板（保留所有原有內容及圖示）
    prs = Presentation(template_path)

    # ===== 第1頁: 封面 - 加入商戶名稱 =====
    if merchant_name and len(prs.slides) >= 1:
        slide1 = prs.slides[0]
        # 在幻燈片底部添加商戶名稱文字框
        merchant_box = slide1.shapes.add_textbox(Inches(0.5), Inches(4.8), Inches(9), Inches(0.6))
        tf = merchant_box.text_frame
        para = tf.paragraphs[0]
        para.text = f"「{merchant_name}」專屬方案"
        para.font.size = Pt(24)
        para.font.bold = True
        para.font.color.rgb = RgbColor(255, 184, 0)  # 金色
        para.alignment = PP_ALIGN.CENTER

    # ===== 第2頁: 痛點與解決方案 - 根據選擇動態生成 =====
    if pain_points and len(prs.slides) >= 2:
        slide2 = prs.slides[1]

        # 查找並替換 "商戶資料【主要業務】和【痛點與解決方案】" 的文字
        for shape in slide2.shapes:
            if shape.has_text_frame:
                if "商戶資料" in shape.text_frame.text or "主要業務" in shape.text_frame.text:
                    # 清空並重寫
                    tf = shape.text_frame
                    tf.clear()
                    para = tf.paragraphs[0]
                    para.text = "商戶資料"
                    para.font.size = Pt(26)
                    para.font.bold = True
                    para.font.color.rgb = RgbColor(34, 34, 34)
                    para.alignment = PP_ALIGN.CENTER

        # 在第2頁添加痛點與解決方案
        # 計算位置：在現有內容下方或適當位置
        y_start = 4.5  # 從幻燈片下半部開始

        for i, pain_key in enumerate(pain_points):
            if pain_key in PAIN_SOLUTIONS:
                data = PAIN_SOLUTIONS[pain_key]

                # 痛點標題
                title_box = slide2.shapes.add_textbox(Inches(0.5), Inches(y_start + i * 0.8), Inches(9), Inches(0.4))
                title_tf = title_box.text_frame
                title_para = title_tf.paragraphs[0]
                title_para.text = f"❌ {data['pain']}"
                title_para.font.size = Pt(14)
                title_para.font.bold = True
                title_para.font.color.rgb = RgbColor(204, 0, 0)

                # 解決方案內容
                sol_box = slide2.shapes.add_textbox(Inches(0.5), Inches(y_start + i * 0.8 + 0.35), Inches(9), Inches(0.5))
                sol_tf = sol_box.text_frame
                sol_para = sol_tf.paragraphs[0]
                sol_para.text = f"✓ {data['solution']}"
                sol_para.font.size = Pt(12)
                sol_para.font.color.rgb = RgbColor(0, 128, 0)

    # 保存
    if output_path:
        prs.save(output_path)
        print(f"PPT 已保存至: {output_path}")
    else:
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        return output

def main():
    parser = argparse.ArgumentParser(description='HKTVmall 商戶招商 PPT 生成器 v3 (保留模板圖示)')
    parser.add_argument('--output', '-o', help='輸出文件路徑')
    parser.add_argument('--template', '-t', help='模板路徑')
    parser.add_argument('--merchant', '-m', help='商戶名稱')
    parser.add_argument('--pains', '-p', help='痛點（逗號分隔）')
    parser.add_argument('--serve', action='store_true', help='啟動 Flask API 服務')
    parser.add_argument('--host', default='0.0.0.0', help='API 主機')
    parser.add_argument('--port', type=int, default=5000, help='API 端口')

    args = parser.parse_args()

    pain_points = []
    if args.pains:
        pain_points = [p.strip() for p in args.pains.split(',')]

    output_path = args.output or f"HKTVmall_招商方案_{args.merchant or '示例'}.pptx"

    print(f"生成 PPT: {output_path}")
    print(f"  商戶: {args.merchant or '示例'}")
    print(f"  痛點: {pain_points or '無'}")
    print(f"  模板: {args.template or DEFAULT_TEMPLATE}")

    try:
        generate_merchant_ppt(
            merchant_name=args.merchant,
            pain_points=pain_points,
            output_path=output_path,
            template_path=args.template
        )
    except FileNotFoundError as e:
        print(f"錯誤: {e}")
        print(f"\n請確認模板檔案存在於: {DEFAULT_TEMPLATE}")

if __name__ == '__main__':
    main()
