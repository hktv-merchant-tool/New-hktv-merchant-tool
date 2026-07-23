#!/usr/bin/env python3
"""
HKTVmall 商戶招商 PPT 生成器
基於 2026 HKTVmall New Merchant Acquisition Deck 模板

使用方法:
  python hktv_ppt_generator.py                    # 生成示例 PPT
  python hktv_ppt_generator.py --serve            # 啟動 Flask API 服務
  python hktv_ppt_generator.py --merchant "ABC電器" --category "本地" --pains "物流成本高,品牌知名度"

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

# ============================================================
# 模板數據 - 基於 2026 HKTVmall New Merchant Acquisition Deck
# ============================================================

# 招商計劃詳情
PLANS = {
    "本地": {
        "name": "Type A Plan - 本地商戶",
        "annual_fee": "HK$25,000 + HK$22,000 廣告金",
        "description": "本地常規商戶 · 年費 HK$25,000 + HK$22,000 廣告金",
        "commission_info": "根據產品類別計算佣金",
        "features": [
            "開立帳戶 + 建立電子合約 + 開展業務",
            "3-5個工作天成功開店",
            "商戶管理平台全面掌控業務營運",
            "全渠道推廣工具",
            "專業客戶服務支援"
        ]
    },
    "海外": {
        "name": "Overseas Plan - 海外商戶",
        "annual_fee": "HK$3,000（原 HK$25,000）",
        "description": "海外跨境商戶 · 年費優惠 HK$3,000",
        "commission_info": "海外跨境商戶專屬佣金率",
        "features": [
            "拓展國際市場客戶",
            "高效無縫的配送解決方案",
            "協助處理海外物流",
            "多語言客戶服務支援",
            "海外市場推廣支援"
        ]
    },
    "服務": {
        "name": "Service Deal - 服務類商戶",
        "annual_fee": "HK$3,000",
        "description": "服務類商戶 · 年費 HK$3,000 · e-Voucher 核銷",
        "commission_info": "服務類商戶專屬佣金率",
        "features": [
            "e-Voucher 核銷系統",
            "O2O 門店網絡",
            "專業服務平台",
            "直播頻道推廣機會",
            "會員引流支援"
        ]
    }
}

# 痛點與解決方案
PAIN_SOLUTIONS = {
    "物流成本高": {
        "pain": "物流成本高",
        "solution": "HKTVmall 提供全港最大住宅配送網絡，500+ 取貨點，350+ 輛貨車，有效降低物流成本。"
    },
    "平台費用不透明": {
        "pain": "平台費用不透明",
        "solution": "清晰的佣金結構，根據產品類別計算，款項每月結算，15個工作日內轉至商戶帳戶。"
    },
    "缺乏電商經驗": {
        "pain": "缺乏電商經驗",
        "solution": "HKTVmall 電商學院提供專業培訓，包括市場推廣技巧、開發市場機會、行內人士分享。"
    },
    "擔心庫存風險": {
        "pain": "擔心庫存風險",
        "solution": "GREEN LAB 臨期百貨計劃，幫助清理庫存；3PL 倉儲服務提供靈活的存貨管理。"
    },
    "希望提升品牌知名度": {
        "pain": "希望提升品牌知名度",
        "solution": "超過50種客制化廣告模板，LEGO、蘇寧等品牌成功案例，TVC廣告推廣機會。"
    },
    "希望拓展年輕客群": {
        "pain": "希望拓展年輕客群",
        "solution": "全面的客戶數據分析，精準CRM定位，每季度4.7x購買頻率的忠實客戶群。"
    },
    "希望增加線上曝光": {
        "pain": "希望增加線上曝光",
        "solution": "全年過億重磅廣告投放，覆蓋58個MTR站，23列全車身包裝列車，數碼引流合作計劃。"
    }
}

# 平台核心數據
PLATFORM_STATS = {
    "gmv_2025": "HK$7.89 Billion",
    "daily_orders": "10萬+",
    "logistics_network": "350+ 輛貨車",
    "o2o_stores": "75家",
    "pickup_points": "500+",
    "annual_ad_spend": "過億",
    "mtr_coverage": "58個MTR站",
    "train_coverage": "23列全車身",
    "customer_purchase_freq": "每季度4.7x"
}

# ============================================================
# PPT 幻燈片內容
# ============================================================

def create_title_slide(prs, merchant_name=None):
    """封面"""
    slide_layout = prs.slide_layouts[6]  # 空白佈局
    slide = prs.slides.add_slide(slide_layout)

    # 標題
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "香港網上購物平台"
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.alignment = PP_ALIGN.CENTER

    # 副標題
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.8), Inches(9), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    sub_para = subtitle_frame.paragraphs[0]
    sub_para.text = "No.1 商戶合作方案"
    sub_para.font.size = Pt(36)
    sub_para.alignment = PP_ALIGN.CENTER

    # 商戶名稱（如果提供）
    if merchant_name:
        merchant_box = slide.shapes.add_textbox(Inches(0.5), Inches(5), Inches(9), Inches(0.8))
        merchant_frame = merchant_box.text_frame
        merchant_para = merchant_frame.paragraphs[0]
        merchant_para.text = f"「{merchant_name}」專屬方案"
        merchant_para.font.size = Pt(28)
        merchant_para.font.color.rgb = RgbColor(0, 102, 204)
        merchant_para.alignment = PP_ALIGN.CENTER

    return slide

def create_company_intro_slide(prs):
    """公司介紹"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    # 標題
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "HKTVmall 里程碑"
    title_para.font.size = Pt(32)
    title_para.font.bold = True

    # 歷程
    milestones = [
        ("1992", "城市電訊（香港）有限公司成立"),
        ("1997", "香港電視面世"),
        ("2012", "電商正式啟動"),
        ("2015-2016", "香港電視正式轉型為網上購物"),
        ("2017-2018", "物流中心自動化揀貨及倉存系統全面運作"),
        ("2019-2020", "海外業務擴張及開拓業務新板塊"),
        ("2021-2022", "街市即日餸、直播頻道推出"),
        ("2023-現在", "持續引領香港電商市場")
    ]

    y_pos = 1.5
    for year, desc in milestones:
        year_box = slide.shapes.add_textbox(Inches(0.8), Inches(y_pos), Inches(1.5), Inches(0.5))
        year_frame = year_box.text_frame
        year_para = year_frame.paragraphs[0]
        year_para.text = year
        year_para.font.size = Pt(14)
        year_para.font.bold = True

        desc_box = slide.shapes.add_textbox(Inches(2.5), Inches(y_pos), Inches(6.5), Inches(0.5))
        desc_frame = desc_box.text_frame
        desc_para = desc_frame.paragraphs[0]
        desc_para.text = desc
        desc_para.font.size = Pt(14)

        y_pos += 0.55

def create_platform_advantages_slide(prs):
    """平台優勢"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    # 標題
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "市場領先優勢"
    title_para.font.size = Pt(32)
    title_para.font.bold = True

    advantages = [
        "24小時網上購物",
        "物流 & O2O 門市",
        "營銷與廣告",
        "平台大數據分析",
        "發展海外市場",
        "商客互聯平台與直播頻道"
    ]

    y_pos = 1.8
    for i, adv in enumerate(advantages, 1):
        bullet_box = slide.shapes.add_textbox(Inches(1), Inches(y_pos), Inches(8), Inches(0.6))
        bullet_frame = bullet_box.text_frame
        bullet_para = bullet_frame.paragraphs[0]
        bullet_para.text = f"✓ {adv}"
        bullet_para.font.size = Pt(24)
        y_pos += 0.7

def create_traffic_slide(prs):
    """流量與表現"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    # 標題
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "強大流量與卓越表現"
    title_para.font.size = Pt(32)
    title_para.font.bold = True

    # GMV
    gmv_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1))
    gmv_frame = gmv_box.text_frame
    gmv_para = gmv_frame.paragraphs[0]
    gmv_para.text = f"2025 Total GMV: {PLATFORM_STATS['gmv_2025']}"
    gmv_para.font.size = Pt(36)
    gmv_para.font.bold = True
    gmv_para.font.color.rgb = RgbColor(0, 102, 204)

    # 關鍵數據
    stats = [
        (f"每日訂單處理: {PLATFORM_STATS['daily_orders']}", 3.5),
        (f"物流車隊: {PLATFORM_STATS['logistics_network']}", 4.3),
        (f"O2O 門店: {PLATFORM_STATS['o2o_stores']}", 5.1),
        (f"取貨點: {PLATFORM_STATS['pickup_points']}", 5.9)
    ]

    for text, y in stats:
        stat_box = slide.shapes.add_textbox(Inches(1), Inches(y), Inches(8), Inches(0.6))
        stat_frame = stat_box.text_frame
        stat_para = stat_frame.paragraphs[0]
        stat_para.text = text
        stat_para.font.size = Pt(22)

def create_customer_slide(prs):
    """客戶數據"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "全面滲透，擁有高回購率"
    title_para.font.size = Pt(32)
    title_para.font.bold = True

    points = [
        "客戶性別分佈均衡",
        "客戶橫誇各個年齡層",
        "每季度的客戶平均購買頻率",
        f"每季度{PLATFORM_STATS['customer_purchase_freq']}購買 = 忠實客戶",
        "高重覆購買率 = 為商家提供穩定的銷售"
    ]

    y_pos = 1.8
    for point in points:
        bullet_box = slide.shapes.add_textbox(Inches(1), Inches(y_pos), Inches(8), Inches(0.6))
        bullet_frame = bullet_box.text_frame
        bullet_para = bullet_frame.paragraphs[0]
        bullet_para.text = f"• {point}"
        bullet_para.font.size = Pt(20)
        y_pos += 0.65

def create_logistics_slide(prs):
    """物流系統"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "發展智慧物流"
    title_para.font.size = Pt(32)
    title_para.font.bold = True

    logistics = [
        "機械化自動執貨及倉存系統",
        "自動導引車系統 (AGV)",
        "自動化交叉帶式自動分揀系統",
        "全港最大住宅配送網絡",
        "350+ 輛貨車，每日可處理10萬+訂單量",
        "75家O2O門店"
    ]

    y_pos = 1.8
    for item in logistics:
        bullet_box = slide.shapes.add_textbox(Inches(1), Inches(y_pos), Inches(8), Inches(0.6))
        bullet_frame = bullet_box.text_frame
        bullet_para = bullet_frame.paragraphs[0]
        bullet_para.text = f"• {item}"
        bullet_para.font.size = Pt(20)
        y_pos += 0.65

def create_marketing_slide(prs):
    """營銷工具"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "高效即用推廣工具"
    title_para.font.size = Pt(32)
    title_para.font.bold = True

    marketing = [
        "超過50種客制化廣告模板",
        "圖像、文字廣告、分類全頁廣告",
        "橫幅廣告、開頁廣告、關鍵字廣告",
        "精準客戶群組 (CRM) 定位",
        "行銷自動化工具",
        f"單日銷售額突破 {PLATFORM_STATS['annual_ad_spend']} HKTVmall全額承擔85折"
    ]

    y_pos = 1.8
    for item in marketing:
        bullet_box = slide.shapes.add_textbox(Inches(1), Inches(y_pos), Inches(8), Inches(0.6))
        bullet_frame = bullet_box.text_frame
        bullet_para = bullet_frame.paragraphs[0]
        bullet_para.text = f"• {item}"
        bullet_para.font.size = Pt(20)
        y_pos += 0.65

def create_pain_solution_slide(prs, pain_points):
    """痛點與解決方案"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "HKTVmall 助您解決商業挑戰"
    title_para.font.size = Pt(32)
    title_para.font.bold = True

    y_pos = 1.5
    for pain_key in pain_points:
        if pain_key in PAIN_SOLUTIONS:
            data = PAIN_SOLUTIONS[pain_key]

            # 痛點
            pain_box = slide.shapes.add_textbox(Inches(0.8), Inches(y_pos), Inches(8.5), Inches(0.4))
            pain_frame = pain_box.text_frame
            pain_para = pain_frame.paragraphs[0]
            pain_para.text = f"❌ {data['pain']}"
            pain_para.font.size = Pt(16)
            pain_para.font.bold = True
            pain_para.font.color.rgb = RgbColor(204, 0, 0)

            # 解決方案
            sol_box = slide.shapes.add_textbox(Inches(1.2), Inches(y_pos + 0.35), Inches(8), Inches(0.6))
            sol_frame = sol_box.text_frame
            sol_para = sol_frame.paragraphs[0]
            sol_para.text = f"✓ {data['solution']}"
            sol_para.font.size = Pt(14)
            sol_para.font.color.rgb = RgbColor(0, 128, 0)

            y_pos += 0.95

def create_plan_slide(prs, plan_key):
    """招商計劃詳情"""
    if plan_key not in PLANS:
        return

    plan = PLANS[plan_key]
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    # 標題
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = plan['name']
    title_para.font.size = Pt(32)
    title_para.font.bold = True

    # 年費
    fee_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(0.8))
    fee_frame = fee_box.text_frame
    fee_para = fee_frame.paragraphs[0]
    fee_para.text = f"年費: {plan['annual_fee']}"
    fee_para.font.size = Pt(24)
    fee_para.font.color.rgb = RgbColor(0, 102, 204)

    # 描述
    desc_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.3), Inches(9), Inches(0.6))
    desc_frame = desc_box.text_frame
    desc_para = desc_frame.paragraphs[0]
    desc_para.text = plan['description']
    desc_para.font.size = Pt(18)

    # 特點
    y_pos = 3.2
    features_box = slide.shapes.add_textbox(Inches(0.8), Inches(y_pos), Inches(8), Inches(4))
    features_frame = features_box.text_frame
    for i, feature in enumerate(plan['features']):
        if i == 0:
            p = features_frame.paragraphs[0]
        else:
            p = features_frame.add_paragraph()
        p.text = f"✓ {feature}"
        p.font.size = Pt(18)
        p.space_after = Pt(10)

def create_success_cases_slide(prs):
    """成功案例"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "真實商家成效"
    title_para.font.size = Pt(32)
    title_para.font.bold = True

    cases = [
        ("蘇寧 x HKTVmall 直播", "90萬", "一晚GMV接近，1700件+產品"),
        ("中小企新乳酪品牌", "3萬", "單場直播售出477杯"),
        ("中小企2025入駐新品牌", "1000+件", "第一季已售出"),
        ("小品牌業績增長", "3倍", "年增長率")
    ]

    y_pos = 1.8
    for case_name, revenue, desc in cases:
        # 案例名
        case_box = slide.shapes.add_textbox(Inches(0.8), Inches(y_pos), Inches(8), Inches(0.5))
        case_frame = case_box.text_frame
        case_para = case_frame.paragraphs[0]
        case_para.text = f"【{case_name}】"
        case_para.font.size = Pt(18)
        case_para.font.bold = True

        # 數據
        data_box = slide.shapes.add_textbox(Inches(1.2), Inches(y_pos + 0.4), Inches(7.5), Inches(0.8))
        data_frame = data_box.text_frame
        data_para = data_frame.paragraphs[0]
        data_para.text = f"  {revenue} - {desc}"
        data_para.font.size = Pt(16)

        y_pos += 1.3

def create_joining_slide(prs):
    """簡易加入流程"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "簡易加入流程"
    title_para.font.size = Pt(32)
    title_para.font.bold = True

    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(0.6))
    subtitle_frame = subtitle_box.text_frame
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.text = "3個簡單步驟 至 3-5個工作天成功開店"
    subtitle_para.font.size = Pt(24)
    subtitle_para.font.color.rgb = RgbColor(0, 102, 204)

    steps = [
        ("第一步", "開立帳戶", "準備商業登記證(BR)、銀行月結單"),
        ("第二步", "建立電子合約", "完成合約簽署程序"),
        ("第三步", "開展您的業務", "開始上架產品並接受訂單")
    ]

    y_pos = 2.3
    for step_title, step_name, step_desc in steps:
        # 步驟標題
        st_box = slide.shapes.add_textbox(Inches(0.8), Inches(y_pos), Inches(2), Inches(0.5))
        st_frame = st_box.text_frame
        st_para = st_frame.paragraphs[0]
        st_para.text = step_title
        st_para.font.size = Pt(20)
        st_para.font.bold = True
        st_para.font.color.rgb = RgbColor(0, 102, 204)

        # 步驟名
        sn_box = slide.shapes.add_textbox(Inches(2.8), Inches(y_pos), Inches(3), Inches(0.5))
        sn_frame = sn_box.text_frame
        sn_para = sn_frame.paragraphs[0]
        sn_para.text = step_name
        sn_para.font.size = Pt(20)
        sn_para.font.bold = True

        # 步驟描述
        sd_box = slide.shapes.add_textbox(Inches(5.8), Inches(y_pos), Inches(3.5), Inches(0.5))
        sd_frame = sd_box.text_frame
        sd_para = sd_frame.paragraphs[0]
        sd_para.text = step_desc
        sd_para.font.size = Pt(16)

        y_pos += 0.8

def create_contact_slide(prs, contact_email=None, merchant_name=None):
    """聯絡我們"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    # 標題
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = "立即加入我們"
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.alignment = PP_ALIGN.CENTER

    # 聯絡方式
    email = contact_email or "aog.merc@hktv.com.hk"

    contact_box = slide.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(9), Inches(2.5))
    contact_frame = contact_box.text_frame
    contact_frame.paragraphs[0].text = f"電郵地址: {email}"
    contact_frame.paragraphs[0].font.size = Pt(24)
    contact_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    p2 = contact_frame.add_paragraph()
    p2.text = "網站: http://business.hktvmall.com"
    p2.font.size = Pt(24)
    p2.alignment = PP_ALIGN.CENTER

    p3 = contact_frame.add_paragraph()
    p3.text = "WhatsApp: +852 5283 4138"
    p3.font.size = Pt(24)
    p3.alignment = PP_ALIGN.CENTER

    p4 = contact_frame.add_paragraph()
    p4.text = "https://wa.me/85252834138"
    p4.font.size = Pt(20)
    p4.alignment = PP_ALIGN.CENTER

    if merchant_name:
        p5 = contact_frame.add_paragraph()
        p5.text = f"\n\n「{merchant_name}」期待與您合作！"
        p5.font.size = Pt(22)
        p5.font.color.rgb = RgbColor(0, 102, 204)
        p5.alignment = PP_ALIGN.CENTER

# ============================================================
# 主 PPT 生成函數
# ============================================================

def generate_merchant_ppt(merchant_name=None, category="本地", pain_points=None, contact_email="aog.merc@hktv.com.hk", output_path=None):
    """
    生成商戶招商 PPT

    參數:
        merchant_name: 商戶名稱
        category: "本地" | "海外" | "服務"
        pain_points: 痛點列表
        contact_email: 聯絡電郵
        output_path: 輸出文件路徑
    """
    if pain_points is None:
        pain_points = []

    # 創建演示文稿
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # 1. 封面
    create_title_slide(prs, merchant_name)

    # 2. 公司介紹
    create_company_intro_slide(prs)

    # 3. 平台優勢
    create_platform_advantages_slide(prs)

    # 4. 流量與表現
    create_traffic_slide(prs)

    # 5. 客戶數據
    create_customer_slide(prs)

    # 6. 物流系統
    create_logistics_slide(prs)

    # 7. 營銷工具
    create_marketing_slide(prs)

    # 8. 痛點與解決方案（如果選擇了痛點）
    if pain_points:
        create_pain_solution_slide(prs, pain_points)

    # 9. 成功案例
    create_success_cases_slide(prs)

    # 10. 招商計劃
    create_plan_slide(prs, category)

    # 11. 加入流程
    create_joining_slide(prs)

    # 12. 聯絡我們
    create_contact_slide(prs, contact_email, merchant_name)

    # 保存
    if output_path:
        prs.save(output_path)
        print(f"PPT 已保存至: {output_path}")
    else:
        # 返回 BytesIO
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

        # 生成 PPT
        ppt_file = generate_merchant_ppt(
            merchant_name=merchant_name,
            category=category,
            pain_points=pain_points,
            contact_email=contact_email
        )

        # 返回文件
        return send_file(
            ppt_file,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            as_attachment=True,
            download_name=f'HKTVmall_招商方案_{merchant_name}.pptx'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    """健康檢查"""
    return jsonify({'status': 'ok'})

# ============================================================
# 主程序
# ============================================================

def main():
    parser = argparse.ArgumentParser(description='HKTVmall 商戶招商 PPT 生成器')
    parser.add_argument('--serve', action='store_true', help='啟動 Flask API 服務')
    parser.add_argument('--output', '-o', help='輸出文件路徑')
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
        # 命令列模式生成 PPT
        pain_points = []
        if args.pains:
            pain_points = [p.strip() for p in args.pains.split(',')]

        output_path = args.output or f'HKTVmall_招商方案_{args.merchant or "示例"}.pptx'

        print(f"生成 PPT: {output_path}")
        print(f"  商戶: {args.merchant or '示例'}")
        print(f"  計劃: {args.category}")
        print(f"  痛點: {pain_points or '無'}")

        generate_merchant_ppt(
            merchant_name=args.merchant,
            category=args.category,
            pain_points=pain_points,
            contact_email=args.email,
            output_path=output_path
        )

if __name__ == '__main__':
    main()
