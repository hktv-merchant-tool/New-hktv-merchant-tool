#!/usr/bin/env python3
"""
HKTVmall 商戶招商 PPT 生成器 v4
基於 Short Version 模板 (22頁)

使用方式:
  python hktv_ppt_generator.py --merchant "ABC電器" --pains "物流成本高,平台費用不透明"

指令 1: 頁面 [1, 3-22] 完全保留原有內容（圖示、圖片、形狀）
指令 2: 頁面 [2] 根據痛點選擇動態生成

模板路徑: C:\\Users\\cylee\\Desktop\\2026 HKTVmall New Merchant Acquisition Deck_Short Verson [CHI].pptx
"""

import io
import os
import argparse
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# 痛點與解決方案
PAIN_SOLUTIONS = {
    "物流成本高": {"pain": "物流成本高", "solution": "全港最大住宅配送網絡，350+輛貨車，500+取貨點，大幅降低物流成本"},
    "平台費用不透明": {"pain": "平台費用不透明", "solution": "佣金按分類碼固定比例計算，每月15個工作日結算，無隱藏收費"},
    "缺乏電商經驗": {"pain": "缺乏電商經驗", "solution": "電商學院提供完整培訓，MMS後台操作、產品上架、廣告投放一對一支援"},
    "擔心庫存風險": {"pain": "擔心庫存風險", "solution": "GREEN LAB臨期百貨計劃，3PL靈活存貨管理，代售模式先售後買降低風險"},
    "希望提升品牌知名度": {"pain": "希望提升品牌知名度", "solution": "每年過億廣告投放，TVB黃金時段，58個MTR站，23列全車身，260萬+買家"},
    "希望拓展年輕客群": {"pain": "希望拓展年輕客群", "solution": "精準CRM定位，25-40歲核心消費群，每季度4.7x購買頻率的忠實客戶"},
    "希望增加線上曝光": {"pain": "希望增加線上曝光", "solution": "關鍵字廣告、橫幅廣告、首頁推廣位，SEO及站內搜尋排名優化"}
}

# 預設模板路徑（Windows 路徑）
DEFAULT_TEMPLATE = "C:\\Users\\cylee\\Desktop\\2026 HKTVmall New Merchant Acquisition Deck_Short Verson [CHI].pptx"

# WSL 路徑映射
WSL_TEMPLATE = "/mnt/c/Users/cylee/Desktop/2026 HKTVmall New Merchant Acquisition Deck_Short Verson [CHI].pptx"

def generate_merchant_ppt(merchant_name=None, pain_points=None, output_path=None, template_path=None, merchant_desc=None):
    """
    生成商戶招商 PPT - 完全保留模板原有內容

    指令 1: 頁面 [1, 3-22] 完全保留原有內容
    指令 2: 頁面 [2] 根據痛點選擇動態生成，並添加業務描述
    """
    if pain_points is None:
        pain_points = []

    # 確定模板路徑
    if template_path is None:
        # 嘗試多個可能的路徑
        for p in [WSL_TEMPLATE, DEFAULT_TEMPLATE]:
            if os.path.exists(p):
                template_path = p
                break

    if template_path is None or not os.path.exists(template_path):
        raise FileNotFoundError(f"找不到模板檔案: {template_path}")

    # 開啟模板（完全保留所有原有內容）
    prs = Presentation(template_path)

    # ===== 第1頁: 添加商戶名稱 =====
    # 原頁面內容完全保留，只在標題處添加商戶名稱
    if merchant_name and len(prs.slides) > 0:
        slide1 = prs.slides[0]
        for shape in slide1.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        if "No.1" in run.text:
                            run.text = run.text.replace("No.1", merchant_name)
                            break

    # ===== 第2頁: 添加業務描述及痛點與解決方案 =====
    if len(prs.slides) > 1:
        slide2 = prs.slides[1]
        y = 4.0  # 從幻燈片下半部開始

        # 先添加業務描述（如果有）
        if merchant_desc:
            # 業務描述標題
            desc_title_box = slide2.shapes.add_textbox(Inches(0.5), Inches(y), Inches(9), Inches(0.4))
            desc_tf = desc_title_box.text_frame
            desc_para = desc_tf.paragraphs[0]
            desc_para.text = f"【業務描述】{merchant_desc}"
            desc_para.font.size = Pt(13)
            desc_para.font.bold = True
            desc_para.font.color.rgb = RGBColor(0, 102, 204)
            y += 0.55

        # 添加痛點與解決方案
        if pain_points:
            # 痛點標題
            pain_header_box = slide2.shapes.add_textbox(Inches(0.5), Inches(y), Inches(9), Inches(0.4))
            pain_header_tf = pain_header_box.text_frame
            pain_header_para = pain_header_tf.paragraphs[0]
            pain_header_para.text = "【痛點與解決方案】"
            pain_header_para.font.size = Pt(14)
            pain_header_para.font.bold = True
            pain_header_para.font.color.rgb = RGBColor(102, 102, 102)
            y += 0.45

            for pain_key in pain_points:
                if pain_key in PAIN_SOLUTIONS:
                    data = PAIN_SOLUTIONS[pain_key]
                    # 痛點標題
                    title_box = slide2.shapes.add_textbox(Inches(0.5), Inches(y), Inches(9), Inches(0.4))
                    tf = title_box.text_frame
                    para = tf.paragraphs[0]
                    para.text = f"❌ {data['pain']}"
                    para.font.size = Pt(14)
                    para.font.bold = True
                    para.font.color.rgb = RGBColor(204, 0, 0)

                    # 解決方案內容
                    sol_box = slide2.shapes.add_textbox(Inches(0.5), Inches(y + 0.35), Inches(9), Inches(0.5))
                    sol_tf = sol_box.text_frame
                    sol_para = sol_tf.paragraphs[0]
                    sol_para.text = f"✓ {data['solution']}"
                    sol_para.font.size = Pt(12)
                    sol_para.font.color.rgb = RGBColor(0, 128, 0)

                    y += 0.75

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
    parser = argparse.ArgumentParser(description='HKTVmall 商戶招商 PPT 生成器 v4（保留模板圖示）')
    parser.add_argument('--output', '-o', help='輸出文件路徑')
    parser.add_argument('--template', '-t', help='模板路徑')
    parser.add_argument('--merchant', '-m', help='商戶名稱')
    parser.add_argument('--pains', '-p', help='痛點（逗號分隔）')
    parser.add_argument('--desc', '-d', help='業務描述')
    args = parser.parse_args()

    pain_points = []
    if args.pains:
        pain_points = [p.strip() for p in args.pains.split(',')]

    output_path = args.output or f"HKTVmall_招商方案_{args.merchant or '示例'}.pptx"

    print(f"生成 PPT: {output_path}")
    print(f"  商戶: {args.merchant or '（只用於第1頁）'}")
    print(f"  業務描述: {args.desc or '無'}")
    print(f"  痛點: {pain_points or '無'}")
    print(f"  模板: {args.template or WSL_TEMPLATE}")
    print("  頁面 1, 3-22: 完全保留原有內容")

    try:
        generate_merchant_ppt(
            merchant_name=args.merchant,
            pain_points=pain_points,
            output_path=output_path,
            template_path=args.template,
            merchant_desc=args.desc
        )
    except FileNotFoundError as e:
        print(f"\n錯誤: {e}")
        print(f"\n請確認模板存在: {WSL_TEMPLATE}")

if __name__ == '__main__':
    main()
