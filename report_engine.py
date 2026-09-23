"""
report_engine.py
Core logic extracted from Enhancement_of_Report_Automation_v20260922.ipynb.
All functions are preserved exactly as in the notebook.
"""

import os
import sys
import platform
import subprocess
import smtplib
import json
import re
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import io

import pandas as pd
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend for server use
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

# ─── Color Palette Constants ──────────────────────────────────────────────────

# Dark Theme
COLOR_BG_DARK       = RGBColor(30,  41,  59)
COLOR_CARD_DARK     = RGBColor(15,  23,  42)
COLOR_TEXT_LIGHT    = RGBColor(241, 245, 249)
COLOR_TEXT_MUTED    = RGBColor(148, 163, 184)
COLOR_ACCENT_BLUE   = RGBColor(59,  130, 246)

# Light Theme
COLOR_BG_LIGHT      = RGBColor(248, 250, 252)
COLOR_CARD_LIGHT    = RGBColor(255, 255, 255)
COLOR_TEXT_DARK     = RGBColor(15,  23,  42)
COLOR_ACCENT_POWERBI = RGBColor(14, 165, 233)

# Status Colors
COLOR_GREEN = RGBColor(34,  197, 94)
COLOR_RED   = RGBColor(239, 68,  68)


# ─── Load Data ────────────────────────────────────────────────────────────────

def load_and_process_car_sales_data(url_or_path, progress_cb=None):
    """
    Membaca dataset dari file lokal atau URL GitHub/raw.
    Mengembalikan DataFrame mentah.
    progress_cb: optional callable(message: str) untuk streaming log ke UI.
    """
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        else:
            print(msg)

    # Jika URL GitHub biasa, konversi ke raw URL
    raw_url = url_or_path
    if "github.com" in url_or_path and "/blob/" in url_or_path:
        raw_url = url_or_path.replace("github.com", "raw.githubusercontent.com").replace("/blob/", "/")

    log(f"Membaca data dari: {raw_url}")
    try:
        df = pd.read_csv(raw_url)
        log(f"[SUCCESS] Data berhasil dimuat! Total baris: {len(df):,}")
    except Exception as e:
        log(f"[ERROR] Gagal membaca data: {e}")
        raise e

    df.columns = df.columns.str.strip()
    return df


# ─── Table Data Preparation ──────────────────────────────────────────────────

def prepare_table_data(df):
    """
    Menghitung metrics untuk Slide Option 2 (Tabel Performa KPI).
    Kolom yang digunakan: Sale_Date, Sale_ID, Final_Sale_Price_LKR,
                           Car_Brand, Car_Model, Discount_Rate
    """
    df = df.copy()
    df['Sale_Date'] = pd.to_datetime(df['Sale_Date'])

    min_date = df['Sale_Date'].min()
    max_date = df['Sale_Date'].max()
    periode_str = f"{min_date.strftime('%B %Y')} - {max_date.strftime('%B %Y')}"

    total_sales_unit  = df['Sale_ID'].count()
    total_sales_price = df['Final_Sale_Price_LKR'].sum()
    total_car_brand   = df['Car_Brand'].nunique()
    total_car_model   = df['Car_Model'].nunique()
    avg_final_price   = df['Final_Sale_Price_LKR'].mean()
    avg_discount_rate = df['Discount_Rate'].mean()

    formatted_discount = (
        f"{avg_discount_rate * 100:.2f}%"
        if avg_discount_rate <= 1
        else f"{avg_discount_rate:.2f}%"
    )

    return {
        "periode": periode_str,
        "metrics": [
            (1, "Total Sales Unit",             f"{total_sales_unit:,}"),
            (2, "Total Sales Price",            f"${total_sales_price:,.2f}"),
            (3, "Total Car Brand Sold",         f"{total_car_brand:,}"),
            (4, "Total Car Model Sold",         f"{total_car_model:,}"),
            (5, "Average Final Price per Car",  f"${avg_final_price:,.2f}"),
            (6, "Average Discount Rate per Car", formatted_discount),
        ]
    }


def prepare_table_data_option2(df):
    """
    Menghitung metrics lengkap untuk Slide Option 3 (Dashboard Power BI Style).
    """
    df = df.copy()
    df['Sale_Date'] = pd.to_datetime(df['Sale_Date'])

    total_sales_unit    = df['Sale_ID'].count()
    total_sales_price   = df['Final_Sale_Price_LKR'].sum()
    total_car_brand     = df['Car_Brand'].nunique()
    total_car_model     = df['Car_Model'].nunique()
    total_vehicle_year  = df['Vehicle_Year'].nunique()
    total_sales_person  = df['Salesperson_ID'].nunique()

    def fmt_num(val):
        if val >= 1e12:
            return f"{val/1e12:.1f}tn"
        elif val >= 1e9:
            v = val / 1e9
            return f"{v:.0f}bn" if abs(v - round(v)) < 0.05 else f"{v:.1f}bn"
        elif val >= 1e3:
            v = val / 1e3
            return f"{v:.0f}K" if abs(v - round(v)) < 0.05 else f"{v:.1f}K"
        return str(val)

    df['YearMonth'] = df['Sale_Date'].dt.to_period('M')
    monthly_trend = df.groupby('YearMonth')['Sale_ID'].count().reset_index()
    monthly_trend['MonthStr'] = monthly_trend['YearMonth'].dt.strftime('%b %Y')

    top_brands = df.groupby('Car_Brand')['Sale_ID'].count().nlargest(3).reset_index()
    top_models = df.groupby('Car_Model')['Sale_ID'].count().nlargest(3).reset_index()
    transmission = df.groupby('Transmission')['Sale_ID'].count().reset_index()
    payment = df.groupby('Payment_Method')['Sale_ID'].count().reset_index()

    disc_price = df.groupby('Car_Model').agg(
        Avg_Discount=('Discount_Rate', 'mean'),
        Avg_Final_Price=('Final_Sale_Price_LKR', 'mean')
    ).reset_index().sort_values(by='Car_Model')

    return {
        'kpis': [
            ("Total Sales Unit",        fmt_num(total_sales_unit)),
            ("Total Sales Price",       fmt_num(total_sales_price)),
            ("Total Car Brand",         str(total_car_brand)),
            ("Total Car Model",         str(total_car_model)),
            ("Total Car Vehicle Year",  str(total_vehicle_year)),
            ("Total Sales Person",      str(total_sales_person)),
        ],
        'monthly_trend': monthly_trend,
        'top_brands':    top_brands,
        'top_models':    top_models,
        'transmission':  transmission,
        'payment':       payment,
        'disc_price':    disc_price,
    }


# ─── Slide Builders ──────────────────────────────────────────────────────────

def add_table_slide(prs, table_data):
    """
    Option 2: Slide Tabel Performa KPI (clean table style).
    """
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)

    # Title & Subtitle
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.0), Inches(1.2))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

    p1 = tf.paragraphs[0]
    p1.text = "Car Sales Performance"
    p1.font.name = "Arial"
    p1.font.size = Pt(24)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(0, 0, 0)
    p1.space_after = Pt(4)

    p2 = tf.add_paragraph()
    p2.text = f"Performa selama periode bulan {table_data['periode']}"
    p2.font.name = "Arial"
    p2.font.size = Pt(16)
    p2.font.color.rgb = RGBColor(50, 50, 50)

    # Table
    rows = len(table_data['metrics']) + 1
    cols = 3
    table_shape = slide.shapes.add_table(rows, cols, Inches(0.8), Inches(2.0), Inches(11.7), Inches(4.5))
    table = table_shape.table

    table.columns[0].width = Inches(0.8)
    table.columns[1].width = Inches(6.4)
    table.columns[2].width = Inches(4.5)

    # Header
    headers = ["#", "Key Performance Metrics", "Achievement"]
    header_bg = RGBColor(14, 154, 206)
    for col_idx, text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.text = text
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_bg
        p = cell.text_frame.paragraphs[0]
        p.font.name = "Arial"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

    # Data rows
    row_bg_1 = RGBColor(218, 233, 245)
    row_bg_2 = RGBColor(235, 243, 250)
    for row_idx, data_tuple in enumerate(table_data['metrics'], start=1):
        bg_color = row_bg_1 if row_idx % 2 != 0 else row_bg_2
        for col_idx, val in enumerate(data_tuple):
            cell = table.cell(row_idx, col_idx)
            cell.text = str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_color
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Arial"
            p.font.size = Pt(12)
            p.font.color.rgb = RGBColor(30, 30, 30)
            p.alignment = PP_ALIGN.CENTER if col_idx == 0 else PP_ALIGN.LEFT

    return slide


def create_slide_option2(prs, table_data):
    """
    Option 3: Slide Dashboard Power BI Style (KPI Cards + Charts).
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    HEADER_BLUE = RGBColor(106, 127, 193)
    TEXT_DARK   = RGBColor(35,  35,  35)
    CARD_BG     = RGBColor(250, 250, 250)
    CARD_BORDER = RGBColor(230, 230, 230)
    BAR_BLUE    = RGBColor(24,  144, 255)

    def add_card_header(left, top, width, height, title):
        card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER

        header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.35))
        header.fill.solid()
        header.fill.fore_color.rgb = HEADER_BLUE
        header.line.fill.background()

        tf = header.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Segoe UI"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER
        return card

    # Slide Title
    title_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.15), Inches(10), Inches(0.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Car Sales Performance"
    p.font.name = "Segoe UI"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK

    # KPI Cards
    kpi_width  = Inches(1.95)
    kpi_height = Inches(1.25)
    kpi_top    = Inches(0.75)
    kpi_gap    = Inches(0.10)
    start_left = Inches(0.4)

    for i, (title, val) in enumerate(table_data['kpis']):
        left = start_left + i * (kpi_width + kpi_gap)
        add_card_header(left, kpi_top, kpi_width, kpi_height, title)
        val_box = slide.shapes.add_textbox(left, kpi_top + Inches(0.4), kpi_width, Inches(0.8))
        tf_v = val_box.text_frame
        p_v = tf_v.paragraphs[0]
        p_v.text = str(val)
        p_v.font.name = "Segoe UI"
        p_v.font.size = Pt(32)
        p_v.font.bold = True
        p_v.font.color.rgb = TEXT_DARK
        p_v.alignment = PP_ALIGN.CENTER

    # Monthly Trend Chart
    trend_top    = Inches(2.15)
    trend_width  = Inches(12.2)
    trend_height = Inches(2.35)
    add_card_header(start_left, trend_top, trend_width, trend_height, "Monthly Sales Trend")

    cdata = CategoryChartData()
    df_trend = table_data['monthly_trend']
    cdata.categories = df_trend['MonthStr'].tolist()
    cdata.add_series('', df_trend['Sale_ID'].tolist())

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        start_left + Inches(0.1), trend_top + Inches(0.4),
        trend_width - Inches(0.2), trend_height - Inches(0.45),
        cdata
    )
    chart = chart_shape.chart
    chart.has_title = False
    chart.has_legend = False
    series = chart.series[0]
    series.format.line.color.rgb = BAR_BLUE
    series.format.line.width = Pt(2.5)
    chart.category_axis.tick_labels.font.size = Pt(7.5)
    chart.category_axis.tick_labels.font.name = "Segoe UI"
    chart.value_axis.tick_labels.font.size = Pt(7.5)
    chart.value_axis.tick_labels.font.name = "Segoe UI"

    # Bottom row
    bot_top    = Inches(4.60)
    bot_height = Inches(2.65)

    # Top 3 Car Brand
    brand_width = Inches(1.85)
    add_card_header(start_left, bot_top, brand_width, bot_height, "Top 3 Sales by Car Brand")
    cdata_b = CategoryChartData()
    df_b = table_data['top_brands']
    cdata_b.categories = df_b['Car_Brand'].tolist()
    cdata_b.add_series('', df_b['Sale_ID'].tolist())
    chart_b = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        start_left + Inches(0.05), bot_top + Inches(0.4),
        brand_width - Inches(0.1), bot_height - Inches(0.45),
        cdata_b
    ).chart
    chart_b.has_title = False
    chart_b.has_legend = False
    chart_b.value_axis.visible = False
    chart_b.value_axis.has_major_gridlines = False
    series_b = chart_b.series[0]
    series_b.format.fill.solid()
    series_b.format.fill.fore_color.rgb = BAR_BLUE
    plots_b = chart_b.plots[0]
    plots_b.has_data_labels = True
    plots_b.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    plots_b.data_labels.font.size = Pt(8.5)
    plots_b.data_labels.font.name = "Segoe UI"
    plots_b.data_labels.font.bold = True

    # Top 3 Car Model
    model_left  = start_left + brand_width + Inches(0.12)
    model_width = Inches(1.85)
    add_card_header(model_left, bot_top, model_width, bot_height, "Top 3 Sales by Car Model")
    cdata_m = CategoryChartData()
    df_m = table_data['top_models']
    cdata_m.categories = df_m['Car_Model'].tolist()
    cdata_m.add_series('', df_m['Sale_ID'].tolist())
    chart_m = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        model_left + Inches(0.05), bot_top + Inches(0.4),
        model_width - Inches(0.1), bot_height - Inches(0.45),
        cdata_m
    ).chart
    chart_m.has_title = False
    chart_m.has_legend = False
    chart_m.value_axis.visible = False
    chart_m.value_axis.has_major_gridlines = False
    series_m = chart_m.series[0]
    series_m.format.fill.solid()
    series_m.format.fill.fore_color.rgb = BAR_BLUE
    plots_m = chart_m.plots[0]
    plots_m.has_data_labels = True
    plots_m.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    plots_m.data_labels.font.size = Pt(8.5)
    plots_m.data_labels.font.name = "Segoe UI"
    plots_m.data_labels.font.bold = True

    # Transmission Doughnut
    trans_left  = model_left + model_width + Inches(0.12)
    trans_width = Inches(2.05)
    add_card_header(trans_left, bot_top, trans_width, bot_height, "Transmission Type")
    cdata_t = CategoryChartData()
    df_t = table_data['transmission']
    cdata_t.categories = df_t['Transmission'].tolist()
    cdata_t.add_series('', df_t['Sale_ID'].tolist())
    chart_t = slide.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT,
        trans_left + Inches(0.05), bot_top + Inches(0.4),
        trans_width - Inches(0.1), bot_height - Inches(0.45),
        cdata_t
    ).chart
    chart_t.has_title = False
    chart_t.has_legend = True
    chart_t.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart_t.legend.include_in_layout = False
    plots_t = chart_t.plots[0]
    plots_t.has_data_labels = True
    plots_t.data_labels.number_format = '0.0%'
    plots_t.data_labels.show_percentage = True
    plots_t.data_labels.show_value = False
    plots_t.data_labels.font.size = Pt(8.5)
    plots_t.data_labels.font.name = "Segoe UI"
    plots_t.data_labels.font.bold = True

    # Payment Pie
    pay_left  = trans_left + trans_width + Inches(0.12)
    pay_width = Inches(2.20)
    add_card_header(pay_left, bot_top, pay_width, bot_height, "Payment Type")
    cdata_p = CategoryChartData()
    df_p = table_data['payment']
    cdata_p.categories = df_p['Payment_Method'].tolist()
    cdata_p.add_series('', df_p['Sale_ID'].tolist())
    chart_p = slide.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        pay_left + Inches(0.05), bot_top + Inches(0.4),
        pay_width - Inches(0.1), bot_height - Inches(0.45),
        cdata_p
    ).chart
    chart_p.has_title = False
    chart_p.has_legend = True
    chart_p.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart_p.legend.include_in_layout = False
    plots_p = chart_p.plots[0]
    plots_p.has_data_labels = True
    plots_p.data_labels.number_format = '0.0%'
    plots_p.data_labels.show_percentage = True
    plots_p.data_labels.show_value = False
    plots_p.data_labels.font.size = Pt(8.5)
    plots_p.data_labels.font.name = "Segoe UI"
    plots_p.data_labels.font.bold = True

    # Discount & Final Price Table
    table_left  = pay_left + pay_width + Inches(0.12)
    table_width = Inches(3.85)
    add_card_header(table_left, bot_top, table_width, bot_height, "Discount and Final Price")

    df_dp = table_data['disc_price']
    rows = min(len(df_dp) + 1, 7)
    t_shape = slide.shapes.add_table(
        rows, 3,
        table_left + Inches(0.1), bot_top + Inches(0.45),
        table_width - Inches(0.2), Inches(2.0)
    )
    tbl = t_shape.table
    tbl.columns[0].width = Inches(1.25)
    tbl.columns[1].width = Inches(1.10)
    tbl.columns[2].width = Inches(1.30)

    for j, h in enumerate(["Car_Model", "Avg Discount", "Avg Final Price"]):
        cell = tbl.cell(0, j)
        cell.text = h
        p = cell.text_frame.paragraphs[0]
        p.font.name = "Segoe UI"
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = TEXT_DARK
        p.alignment = PP_ALIGN.RIGHT if j > 0 else PP_ALIGN.LEFT

    for i in range(1, rows):
        row_data = df_dp.iloc[i - 1]
        vals = [
            f"⊞ {row_data['Car_Model']}",
            f"{row_data['Avg_Discount']:.2f}",
            f"{row_data['Avg_Final_Price']:,.2f}",
        ]
        bg_col = RGBColor(250, 250, 250) if i % 2 != 0 else RGBColor(240, 243, 248)
        for j, v in enumerate(vals):
            cell = tbl.cell(i, j)
            cell.text = v
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_col
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Segoe UI"
            p.font.size = Pt(9)
            p.font.color.rgb = TEXT_DARK
            p.alignment = PP_ALIGN.RIGHT if j > 0 else PP_ALIGN.LEFT

    return slide


def generate_gemini_dynamic_json(df_raw, api_key, progress_cb=None):
    """
    Mengirimkan metadata dataset ke Gemini API untuk mendapatkan
    Executive Summary & Rekomendasi Visualisasi Tren dalam format JSON.
    """
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        else:
            print(msg)

    if not GENAI_AVAILABLE:
        log("[WARNING] google-genai tidak terinstall, menggunakan fallback.")
        return _gemini_fallback(df_raw)

    client = genai.Client(api_key=api_key)

    num_cols  = df_raw.select_dtypes(include=['number']).columns.tolist()
    date_cols = df_raw.select_dtypes(include=['datetime', 'datetimetz']).columns.tolist()
    if not date_cols:
        for col in df_raw.select_dtypes(include=['object']).columns:
            if any(k in col.lower() for k in ['date', 'tgl', 'time', 'bulan', 'month', 'tahun', 'year']):
                date_cols.append(col)
                break

    total_rows   = len(df_raw)
    summary_dict = {"total_records": total_rows, "numeric_columns": {}}
    for col in num_cols[:5]:
        summary_dict["numeric_columns"][col] = {
            "sum":  float(df_raw[col].sum()),
            "mean": float(df_raw[col].mean()),
        }

    date_col_name = date_cols[0] if date_cols else "N/A"

    prompt = f"""
    Anda adalah seorang Business Analyst Senior. Analisis metadata dataset berikut dan hasilkan respon HANYA dalam format JSON valid (tanpa teks penjelasan lain).

    Data Metadata:
    - Total Baris/Records: {total_rows}
    - Kolom Tanggal/Waktu Terdeteksi: {date_col_name}
    - Ringkasan Kolom Angka: {json.dumps(summary_dict)}

    Format JSON Output yang WAJIB dipatuhi:
    {{
      "title": "Judul Slide Analisis Performa",
      "executive_summary": [
        "- Poin analisis 1 mengenai total performa.",
        "- Poin analisis 2 mengenai tren atau pola utama.",
        "- Poin analisis 3 mengenai saran/rekomendasi strategis."
      ],
      "chart_config": {{
        "chart_title": "Nama Judul Tren Grafik",
        "date_column": "{date_col_name}",
        "value_column": "{num_cols[0] if num_cols else 'N/A'}"
      }}
    }}
    """

    candidate_models = ['gemini-2.0-flash', 'gemini-3.6-flash']
    max_retries = 3

    for model_name in candidate_models:
        for attempt in range(1, max_retries + 1):
            try:
                log(f"[INFO] Menghubungi Gemini API (Model: {model_name}, Attempt {attempt})...")
                response = client.models.generate_content(model=model_name, contents=prompt)
                if response and response.text:
                    text = response.text.strip()
                    if "```json" in text:
                        text = re.search(r"```json(.*?)```", text, re.DOTALL).group(1).strip()
                    elif "```" in text:
                        text = re.search(r"```(.*?)```", text, re.DOTALL).group(1).strip()
                    return json.loads(text)
            except Exception as e:
                safe_err = str(e).encode('ascii', 'ignore').decode('ascii')
                log(f"[WARNING] Kendala pada {model_name}: {safe_err}")
                if "503" in str(e) or "UNAVAILABLE" in str(e):
                    time.sleep(2 ** attempt)
                else:
                    break

    return _gemini_fallback(df_raw)


def _gemini_fallback(df_raw):
    """Fallback JSON jika Gemini API tidak tersedia."""
    num_cols  = df_raw.select_dtypes(include=['number']).columns.tolist()
    date_cols = []
    for col in df_raw.select_dtypes(include=['object']).columns:
        if any(k in col.lower() for k in ['date', 'tgl', 'time', 'bulan', 'month', 'tahun', 'year']):
            date_cols.append(col)
            break

    return {
        "title": "AI Executive Summary & Performance Trend",
        "executive_summary": [
            "- Laporan Performa: Penjualan dan transaksi secara umum menunjukkan kinerja stabil.",
            "- Catatan Tren: Terjadi pertumbuhan berkesinambungan pada beberapa periode utama.",
            "- Rekomendasi: Pertahankan efisiensi operasional pada cabang dengan performa tertinggi.",
        ],
        "chart_config": {
            "chart_title": "Key Metrics Monthly Trend",
            "date_column": date_cols[0] if date_cols else "N/A",
            "value_column": num_cols[0] if num_cols else "N/A",
        },
    }


def create_slide_gemini_summary(prs, df_raw, api_key, progress_cb=None):
    """
    Option 1: Slide AI Executive Summary + Dynamic Trend Chart.
    """
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        else:
            print(msg)

    df_calc = df_raw.copy()
    ai_data = generate_gemini_dynamic_json(df_calc, api_key, progress_cb)

    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    HEADER_BLUE = RGBColor(106, 127, 193)
    TEXT_DARK   = RGBColor(30,  41,  59)
    CARD_BG     = RGBColor(248, 250, 252)
    CARD_BORDER = RGBColor(226, 232, 240)
    BAR_BLUE    = RGBColor(24,  144, 255)

    def add_card_header(left, top, width, height, title):
        card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = CARD_BORDER

        header = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.35))
        header.fill.solid()
        header.fill.fore_color.rgb = HEADER_BLUE
        header.line.fill.background()

        tf = header.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Segoe UI"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER
        return card

    # Slide Title
    title_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.15), Inches(12.5), Inches(0.5))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = ai_data.get("title", "AI Executive Summary & Performance Trend")
    p.font.name = "Segoe UI"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = TEXT_DARK

    # Executive Summary Box
    top_pos_y  = Inches(0.75)
    box_width  = Inches(12.533)
    box_height = Inches(2.4)
    add_card_header(Inches(0.4), top_pos_y, box_width, box_height,
                    "Executive Summary & Strategic Insights (Powered by Gemini AI)")

    text_box = slide.shapes.add_textbox(
        Inches(0.5), top_pos_y + Inches(0.4),
        box_width - Inches(0.2), box_height - Inches(0.45)
    )
    tf_text = text_box.text_frame
    tf_text.word_wrap = True
    tf_text.margin_left  = Inches(0.2)
    tf_text.margin_right = Inches(0.2)
    tf_text.margin_top   = Inches(0.1)

    exec_summaries = ai_data.get("executive_summary", [])
    for idx, line in enumerate(exec_summaries):
        p_line = tf_text.paragraphs[0] if idx == 0 else tf_text.add_paragraph()
        p_line.text = str(line)
        p_line.font.name = "Segoe UI"
        p_line.font.size = Pt(11)
        p_line.font.color.rgb = TEXT_DARK
        p_line.space_after = Pt(6)

    # Dynamic Trend Chart
    chart_pos_y  = Inches(3.30)
    chart_height = Inches(3.80)
    chart_cfg   = ai_data.get("chart_config", {})
    chart_title = chart_cfg.get("chart_title", "Key Metrics Performance Trend")
    date_col    = chart_cfg.get("date_column")
    val_col     = chart_cfg.get("value_column")

    add_card_header(Inches(0.4), chart_pos_y, box_width, chart_height, chart_title)

    cdata = CategoryChartData()
    if date_col and date_col in df_calc.columns and date_col != "N/A":
        try:
            df_calc['__Parsed_Date'] = pd.to_datetime(df_calc[date_col], errors='coerce')
            df_calc['__YearMonth']   = df_calc['__Parsed_Date'].dt.to_period('M')
            if val_col and val_col in df_calc.columns and val_col != "N/A":
                monthly_data = df_calc.groupby('__YearMonth')[val_col].sum().reset_index()
                series_vals  = monthly_data[val_col].tolist()
            else:
                monthly_data = df_calc.groupby('__YearMonth').size().reset_index(name='Count')
                series_vals  = monthly_data['Count'].tolist()
            monthly_data['MonthStr'] = monthly_data['__YearMonth'].dt.strftime('%b %Y')
            cdata.categories = monthly_data['MonthStr'].tolist()
            cdata.add_series('Metric Value', series_vals)
        except Exception:
            counts = df_calc.iloc[:, 0].value_counts().head(10)
            cdata.categories = counts.index.astype(str).tolist()
            cdata.add_series('Volume', counts.values.tolist())
    else:
        num_cols_list = df_calc.select_dtypes(include=['number']).columns
        if len(num_cols_list) > 0:
            sums = df_calc[num_cols_list[:8]].sum()
            cdata.categories = sums.index.astype(str).tolist()
            cdata.add_series('Total', sums.values.tolist())
        else:
            cdata.categories = ['Data']
            cdata.add_series('Total', [len(df_calc)])

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS,
        Inches(0.5), chart_pos_y + Inches(0.45),
        box_width - Inches(0.2), chart_height - Inches(0.55),
        cdata
    )
    chart = chart_shape.chart
    chart.has_title  = False
    chart.has_legend = False
    series = chart.series[0]
    series.format.line.color.rgb = BAR_BLUE
    series.format.line.width = Pt(2.5)
    chart.category_axis.tick_labels.font.size = Pt(8.5)
    chart.category_axis.tick_labels.font.name = "Segoe UI"
    chart.value_axis.tick_labels.font.size = Pt(8.5)
    chart.value_axis.tick_labels.font.name = "Segoe UI"

    log("[SUCCESS] Slide AI Executive Summary + Trend Chart berhasil dibuat!")
    return slide


# ─── PPTX → PDF Conversion ───────────────────────────────────────────────────

def convert_pptx_to_pdf(input_pptx, output_pdf, progress_cb=None):
    """
    Konversi PPTX ke PDF.
    Windows: PowerPoint COM (primary), LibreOffice (fallback).
    Linux/Mac: LibreOffice.
    """
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        else:
            print(msg)

    log("Mulai konversi PPTX ke PDF...")
    current_os = platform.system()
    input_abs  = os.path.abspath(input_pptx)
    output_abs = os.path.abspath(output_pdf)

    # Windows COM Method (PowerPoint installed)
    if current_os == "Windows":
        try:
            import comtypes.client
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1
            presentation = powerpoint.Presentations.Open(input_abs)
            presentation.SaveAs(output_abs, 32)
            presentation.Close()
            powerpoint.Quit()
            log(f"[SUCCESS] Konversi PowerPoint (COM) berhasil! PDF: {output_abs}")
            return True
        except Exception as e:
            log(f"[INFO] PowerPoint COM tidak tersedia ({e}), mencoba LibreOffice...")

    # LibreOffice Fallback
    try:
        if current_os == "Windows":
            lo_path = "soffice"
        elif current_os == "Darwin":
            lo_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        else:
            lo_path = "libreoffice"

        cmd = [
            lo_path, "--headless", "--convert-to", "pdf",
            "--outdir", os.path.dirname(output_abs) or ".", input_abs
        ]
        subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        log(f"[SUCCESS] Konversi LibreOffice sukses! PDF: {output_abs}")
        return True
    except Exception as e:
        log(f"[ERROR] LibreOffice juga gagal: {e}")

    return False


# ─── Email Sending ────────────────────────────────────────────────────────────

def send_report_email_multi(sender_email, sender_password, receiver_emails,
                             file_paths, branch_name, progress_cb=None):
    """
    Mengirimkan email laporan ke multiple penerima untuk cabang tertentu.
    """
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        else:
            print(msg)

    if not receiver_emails:
        log(f"[WARNING] Tidak ada email penerima untuk cabang {branch_name}.")
        return False

    log(f"Menyiapkan email untuk cabang {branch_name} -> {', '.join(receiver_emails)}...")
    smtp_server = "smtp.gmail.com"
    smtp_port   = 587

    msg = MIMEMultipart()
    msg['From']    = sender_email
    msg['To']      = ", ".join(receiver_emails)
    msg['Subject'] = f"Laporan Performa Car Sales Dashboard - Cabang {branch_name}"

    body = (
        f"Halo Team Cabang {branch_name},\n\n"
        f"Berikut kami lampirkan Laporan Performa Penjualan Mobil khusus untuk cabang {branch_name}.\n\n"
        f"Salam,\nAutomation Bot"
    )
    msg.attach(MIMEText(body, 'plain'))

    for filepath in file_paths:
        if os.path.exists(filepath):
            with open(filepath, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f"attachment; filename= {os.path.basename(filepath)}")
            msg.attach(part)

    server = None
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_emails, msg.as_string())
        log(f"[SUCCESS] Email cabang {branch_name} sukses terkirim!")
        return True
    except Exception as e:
        log(f"[ERROR] Gagal mengirimkan email cabang {branch_name}. Error: {e}")
        return False
    finally:
        if server:
            try:
                server.quit()
            except Exception:
                pass


# ─── Email List Loader ────────────────────────────────────────────────────────

def load_email_list_from_excel(excel_path):
    """
    Membaca file list_user_email.xlsx.
    Kolom: role, branch_name, email (multi email dipisah ';')
    Returns: list of dict [{"branch_name": ..., "emails": [...]}, ...]
    """
    try:
        df_excel = pd.read_excel(excel_path)
        col_branch = next(
            (col for col in df_excel.columns if 'branch' in col.lower()),
            df_excel.columns[0]
        )
        col_email = next(
            (col for col in df_excel.columns if 'email' in col.lower()),
            df_excel.columns[1]
        )

        result = []
        for _, row in df_excel.iterrows():
            branch     = str(row[col_branch]).strip()
            emails_raw = str(row[col_email]).strip()
            emails     = [
                e.strip() for e in emails_raw.split(";")
                if e.strip() and e.strip().lower() != 'nan'
            ]
            if branch and branch.lower() != 'nan':
                result.append({"branch_name": branch, "emails": emails})

        return result
    except Exception as e:
        print(f"[ERROR] Gagal membaca Excel email list: {e}")
        return []


# ─── Report Generator (Orchestrator) ─────────────────────────────────────────

def generate_report(df_raw, slide_option, output_dir, gemini_api_key=None,
                    branch_name=None, progress_cb=None):
    """
    Orchestrator utama: buat PPTX + PDF berdasarkan pilihan slide.

    slide_option: "1" | "2" | "3" | "4"
    branch_name : jika diisi, filter data per branch sebelum generate.
    Returns: {"pptx": path, "pdf": path|None, "success": bool}
    """
    def log(msg):
        if progress_cb:
            progress_cb(msg)
        else:
            print(msg)

    df = df_raw.copy()

    # Filter per branch jika ada (kecuali jika branch_name == 'ALL')
    if branch_name and branch_name.strip().upper() != "ALL":
        col_branch_ds = next(
            (col for col in df.columns if 'branch' in col.lower() or 'region' in col.lower()),
            None
        )
        if col_branch_ds:
            df_filtered = df[df[col_branch_ds].astype(str).str.upper() == branch_name.upper()]
            if df_filtered.empty:
                log(f"[WARNING] Data untuk cabang '{branch_name}' tidak ditemukan, menggunakan seluruh data.")
            else:
                df = df_filtered

    # Prepare metrics
    table_data_opt1 = prepare_table_data(df)
    table_data_opt2 = prepare_table_data_option2(df)

    # Build presentation
    prs = Presentation()
    prs.slide_width  = Inches(13.333)
    prs.slide_height = Inches(7.5)

    log(f"[INFO] Membuat slide opsi {slide_option}...")

    if slide_option == "1":
        create_slide_gemini_summary(prs, df, gemini_api_key, progress_cb)
    elif slide_option == "2":
        add_table_slide(prs, table_data_opt1)
    elif slide_option == "3":
        create_slide_option2(prs, table_data_opt2)
    elif slide_option == "4":
        create_slide_gemini_summary(prs, df, gemini_api_key, progress_cb)
        add_table_slide(prs, table_data_opt1)
        create_slide_option2(prs, table_data_opt2)
    else:
        add_table_slide(prs, table_data_opt1)

    os.makedirs(output_dir, exist_ok=True)
    clean_name = (branch_name or "All").replace(" ", "_")
    pptx_path  = os.path.join(output_dir, f"Laporan_Performa_{clean_name}.pptx")
    pdf_path   = os.path.join(output_dir, f"Laporan_Performa_{clean_name}.pdf")

    prs.save(pptx_path)
    log(f"[SUCCESS] PPTX disimpan: {pptx_path}")

    pdf_ok = convert_pptx_to_pdf(pptx_path, pdf_path, progress_cb)

    return {
        "pptx":    pptx_path,
        "pdf":     pdf_path if pdf_ok else None,
        "success": True,
    }
