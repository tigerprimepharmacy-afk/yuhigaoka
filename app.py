"""
夕陽ヶ丘苑 請求書CSV生成アプリ
起動: streamlit run app.py
"""

import os
import re
import csv
import io
import unicodedata
import urllib.request
from datetime import date
from calendar import monthrange

import streamlit as st
import pdfplumber

# ==================== ページ設定 ====================
st.set_page_config(
    page_title='夕陽ヶ丘苑 請求書CSV生成',
    page_icon='🏥',
    layout='wide',
)

st.title('🏥 夕陽ヶ丘苑 請求書CSV生成')
st.caption('調剤報酬PDF・会計スプレッドシートから MoneyForward Freee 用CSVを作成します。')

with st.expander('📖 使い方'):
    st.markdown("""
### 手順

**Step 1｜サイドバーを設定する**

| 項目 | 内容 |
|------|------|
| 請求対象年月 | 請求する年・月を入力 |
| 請求日 | 請求書に記載する日付 |
| お支払期限 | 支払い期限の日付 |

**Step 2｜PDFをアップロードする**
- ① 調剤報酬PDF：「患者別月間負担額一覧」のPDF
- ② 口座振替スケジュールPDF：ICC患者がいる場合のみ（任意）

**Step 3｜「🚀 CSV を生成する」をクリック**

**Step 4｜「📥 ダウンロード」をクリック**してMoneyForward Freeeにインポート

---

### 会計区分と請求書の仕様

| 会計区分 | 宛先 | 内容 |
|---------|------|------|
| **まとめて** | 夕陽ヶ丘苑　御中 | 全患者を1枚にまとめ、品目に患者名と日別内訳を記載 |
| **個人** | 患者名　様 | 個別請求書 |
| **ICC** | 患者名　様 | 個別請求書・備考欄に口座振替日を記載 |
""")



# ==================== PDF パース ====================

def build_column_map(words):
    """
    ヘッダー行から列情報を構築する。
    Returns:
        day_x: {day_num: x_center}  (1〜31日の列中心x座標)
        total_x: 合計額列のx中心
        jitai_x: （自費）列のx中心
    """
    # ヘッダー行は top が約 14〜16 付近
    header_words = [w for w in words if 13 <= w['top'] <= 17]
    day_x = {}
    total_x = None
    jitai_x = None

    for w in header_words:
        text = w['text']
        m = re.match(r'^(\d+)日$', text)
        if m:
            day = int(m.group(1))
            day_x[day] = (w['x0'] + w['x1']) / 2
        elif '合計' in text:
            total_x = (w['x0'] + w['x1']) / 2
        elif '自費' in text:
            jitai_x = (w['x0'] + w['x1']) / 2

    return day_x, total_x, jitai_x


def x_to_day(x, day_x_map, tolerance=10):
    """x座標から最も近い日付番号を返す（許容範囲外はNone）。"""
    if not day_x_map:
        return None
    closest_day, closest_dist = min(
        ((d, abs(cx - x)) for d, cx in day_x_map.items()),
        key=lambda kv: kv[1],
    )
    return closest_day if closest_dist <= tolerance else None


def parse_pdf_per_day(pdf_bytes):
    """
    調剤報酬PDFから患者別・日付別の金額を抽出する。

    Returns:
        {患者名: {'days': {day_num: amount}, 'total': int, 'sentei': int, 'naizetsu': int}}
    """
    results = {}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue

            day_x, total_x, jitai_x = build_column_map(words)
            if not day_x:
                continue

            # 選定療養・内税列のx位置を別途取得
            header_words = [w for w in words if 13 <= w['top'] <= 17]
            sentei_x = None
            naizetsu_x = None
            for w in header_words:
                if '選定' in w['text']:
                    sentei_x = (w['x0'] + w['x1']) / 2
                elif '内税' in w['text']:
                    naizetsu_x = (w['x0'] + w['x1']) / 2

            # データ行（ヘッダーより下）
            data_words = [w for w in words if w['top'] > 17]

            # top 座標（小数点以下四捨五入）でグループ化
            rows_by_top = {}
            for w in data_words:
                key = round(w['top'])
                rows_by_top.setdefault(key, []).append(w)

            for top_key in sorted(rows_by_top):
                row_words = sorted(rows_by_top[top_key], key=lambda w: w['x0'])
                if not row_words:
                    continue

                first_text = row_words[0]['text']

                # 点数行・合計行をスキップ
                if '点数' in first_text or '合計' in first_text or '総' in first_text:
                    continue

                # 患者名列（x < 65）の非数値ワード
                name_words = [
                    w for w in row_words
                    if w['x0'] < 65 and not w['text'].replace(',', '').isdigit()
                ]
                if not name_words:
                    continue

                patient_name = ' '.join(w['text'] for w in name_words).strip()
                if not patient_name:
                    continue

                if patient_name not in results:
                    results[patient_name] = {
                        'days': {}, 'total': 0, 'sentei': 0, 'naizetsu': 0
                    }

                # 数値を各列へマッピング
                for w in row_words:
                    if w['x0'] < 65:
                        continue
                    text = w['text'].replace(',', '')
                    if not text.isdigit():
                        continue
                    amount = int(text)
                    if amount == 0:
                        continue

                    x = (w['x0'] + w['x1']) / 2

                    if total_x and abs(x - total_x) < 15:
                        results[patient_name]['total'] = amount
                    elif jitai_x and abs(x - jitai_x) < 15:
                        pass  # 自費はスキップ
                    elif sentei_x and abs(x - sentei_x) < 15:
                        results[patient_name]['sentei'] = amount
                    elif naizetsu_x and abs(x - naizetsu_x) < 15:
                        results[patient_name]['naizetsu'] = amount
                    else:
                        day = x_to_day(x, day_x)
                        if day:
                            results[patient_name]['days'][day] = amount

    return results


# ==================== Google Sheets パース ====================

# 人名でよく使われる異体字 → 標準字体 のマッピング
_VARIANT_MAP = str.maketrans({
    '髙': '高',  # U+9AD9 → U+9AD8
    '濵': '浜',  # U+6FF5 → U+6D5C
    '濱': '浜',  # U+6FF1 → U+6D5C
    '德': '徳',  # U+5FB7 → U+5FB3 (旧字体)
    '眞': '真',  # U+771E → U+771F
    '邉': '辺',  # U+9089 → U+8FBA
    '邊': '辺',  # U+908A → U+8FBA
    '桒': '桑',  # U+6852 → U+sang
    '﨑': '崎',  # U+FA11 (互換漢字) → U+5D0E ※NKFCで対応済みだが念のため
})


def normalize_name(name):
    """
    氏名の表記ゆれを吸収する。
    ・スペース（全角・半角）除去
    ・NFKC正規化（﨑→崎 等の互換漢字を標準化）
    ・異体字マッピング（髙→高 等）
    """
    name = unicodedata.normalize('NFKC', name)
    name = name.translate(_VARIANT_MAP)
    return re.sub(r'\s+', '', name)


def fetch_kaike_sheet(sheets_url, gid=None):
    """
    会計シートから患者名（A列）と会計方法（B列）を取得する。

    Returns:
        {normalize_name(患者名): {'name': str, 'payment': str}}
    """
    # export URL を生成
    base = re.sub(r'/edit.*$', '', sheets_url)
    base = re.sub(r'/view.*$', '', base)
    export_url = base + '/export?format=csv'

    # GID を解決（引数 > URL フラグメント の順）
    if gid is None:
        m = re.search(r'[#&]gid=(\d+)', sheets_url)
        if m:
            gid = m.group(1)

    if gid is not None:
        export_url += f'&gid={gid}'

    req = urllib.request.Request(export_url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req) as resp:
        raw = resp.read().decode('utf-8')

    reader = csv.reader(io.StringIO(raw))
    data = {}
    for row in reader:
        if len(row) < 2:
            continue
        name    = row[0].strip()
        payment = row[1].strip()

        if not name:
            continue
        # 【藤】【桜】等のセクションヘッダー行をスキップ
        if re.match(r'^【.+】', name):
            continue
        # 列ヘッダー行をスキップ
        if name in ('患者名', '氏名', '名前', 'A'):
            continue
        # 有効な会計方法（まとめて/個人/ICC）のみ取り込む
        if payment not in ('まとめて', '個人', 'ICC'):
            continue

        # ※以降の注記（例: ※振込先要記入）を除去して純粋な氏名を取得
        display_name = re.sub(r'[※＊\*].*$', '', name).strip()

        key = normalize_name(display_name)
        data[key] = {'name': display_name, 'payment': payment}

    return data


# ==================== 口座振替スケジュール PDF パース ====================

def parse_transfer_schedule(pdf_bytes):
    """
    石川コンピュータ・センターの口座振替スケジュールPDFを解析し、
    請求月 → 振替日(月, 日) のマッピングを返す。
    """
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = '\n'.join(page.extract_text() or '' for page in pdf.pages)

    text = re.sub(r'll/', '11/', text)
    text = re.sub(r'\bl/', '1/', text)
    text = re.sub(r'/l(?=[^a-zA-Z])', '/1', text)
    text = re.sub(r'(?<=[^\d])l(?=\d)', '1', text)
    text = re.sub(r'\b1\s1/', '11/', text)

    date_re = re.compile(r'(\d{1,2})/(\d{1,2})(?:\s*[（(][月火水木金土日][）)])?')
    schedule = {}

    for line in text.splitlines():
        dates = date_re.findall(line)
        if len(dates) < 6:
            continue
        input_month = int(dates[2][0])
        billing_month = input_month - 1 if input_month > 1 else 12
        transfer_month = int(dates[5][0])
        transfer_day   = int(dates[5][1])
        schedule[billing_month] = (transfer_month, transfer_day)

    return schedule


def transfer_date_text(schedule, billing_month):
    """請求月に対応する振替日テキストを返す（例: '4月22日'）。"""
    if billing_month not in schedule:
        return None
    m, d = schedule[billing_month]
    return f'{m}月{d}日'


# 2026年の口座振替スケジュール（組み込み済み）
# 請求月 → (振替月, 振替日)
SCHEDULE_2026 = {
    1:  (2, 24),   # 1月分  → 2月24日
    2:  (3, 23),   # 2月分  → 3月23日
    3:  (4, 22),   # 3月分  → 4月22日
    4:  (5, 22),   # 4月分  → 5月22日
    5:  (6, 22),   # 5月分  → 6月22日
    6:  (7, 22),   # 6月分  → 7月22日
    7:  (8, 24),   # 7月分  → 8月24日
    8:  (9, 24),   # 8月分  → 9月24日
    9:  (10, 22),  # 9月分  → 10月22日
    10: (11, 24),  # 10月分 → 11月24日
    11: (12, 22),  # 11月分 → 12月22日
    12: (1, 22),   # 12月分 → 翌年1月22日
}


# ==================== CSV 生成 ====================

HEADER_COLS = [
    'csv_type(変更不可)', '行形式', '取引先名称', '件名', '請求日', 'お支払期限',
    '請求書番号', '売上計上日', 'メモ', 'タグ', '小計', '消費税', '合計金額',
    '取引先敬称', '取引先郵便番号', '取引先都道府県', '取引先住所1', '取引先住所2',
    '取引先部署', '取引先担当者役職', '取引先担当者氏名', '自社担当者氏名',
    '備考', '振込先', '入金ステータス', 'メール送信ステータス', '郵送ステータス',
    'ダウンロードステータス',
    '納品日', '品名', '品目コード', '単価', '数量', '単位', '納品書番号',
    '詳細', '金額', '品目消費税率'
]
N = len(HEADER_COLS)


def empty_row():
    return [''] * N


def day_breakdown(days_dict):
    """日別金額を '16日 1,350円、26日 4,670円' 形式にフォーマット。"""
    parts = []
    for day in sorted(days_dict.keys()):
        amount = days_dict[day]
        parts.append(f'{day}日 {amount:,}円')
    return '、'.join(parts)


# PDF患者名（スペースあり） → 正規化キーで payment_data と突き合わせるヘルパー
def find_pdf_entry(patient_name_normalized, pdf_data):
    """
    PDF の患者名（正規化済み）で pdf_data を検索する。
    PDF 側のキーも正規化して比較。
    """
    for pdf_name, val in pdf_data.items():
        if normalize_name(pdf_name) == patient_name_normalized:
            return val
    return None


def build_csv(pdf_data, payment_data, billing_label,
              invoice_date, payment_due, billing_date,
              transfer_text, furikomi_info):
    """
    CSV を生成する。

    Parameters
    ----------
    pdf_data : {患者名(スペースあり): {'days': {day: amount}, 'total': int, ...}}
    payment_data : {normalize_name: {'name': str, 'payment': str}}
    transfer_text : str | None  口座振替日テキスト
    furikomi_info : str  振込先情報

    Returns
    -------
    (csv_bytes, included_list, skipped_list)
    """
    rows_out = [HEADER_COLS]
    included = []
    skipped  = []

    # まとめて患者を一時収集 [(display_name, days_dict, total)]
    matome_patients = []

    # pdf_data の正規化キーを作成（検索高速化）
    pdf_by_norm = {normalize_name(k): v for k, v in pdf_data.items()}

    # payment_data の患者を順番に処理（氏名の五十音順）
    for norm_key, pay_info in sorted(payment_data.items(),
                                     key=lambda x: x[1]['name']):
        display_name = pay_info['name']
        payment      = pay_info['payment']

        # PDF データを検索
        pdf_entry = pdf_by_norm.get(norm_key)
        if pdf_entry is None or pdf_entry['total'] == 0:
            skipped.append({'患者名': display_name,
                            '理由': 'PDF未一致または請求額0円'})
            continue

        total     = pdf_entry['total']
        days      = pdf_entry['days']
        breakdown = day_breakdown(days)

        if payment == 'まとめて':
            matome_patients.append((display_name, days, total))
            continue

        # 個人 / ICC: 個別請求書
        is_icc = (payment == 'ICC')
        biko   = f'{transfer_text}口座振替' if (is_icc and transfer_text) else ''

        r = empty_row()
        r[0]  = '40101'
        r[1]  = '請求書'
        r[2]  = display_name
        r[3]  = f'{billing_label} 薬局自己負担額'
        r[4]  = invoice_date
        r[5]  = payment_due
        r[7]  = billing_date
        r[10] = str(total)
        r[11] = '0'
        r[12] = str(total)
        r[13] = '様'
        r[22] = biko
        r[23] = ' ' if is_icc else furikomi_info
        rows_out.append(r)

        # 品目行
        ri = empty_row()
        ri[0]  = '40101'
        ri[1]  = '品目'
        ri[29] = '調剤報酬（自己負担）'
        ri[31] = str(total)
        ri[32] = '1'
        ri[35] = breakdown
        ri[36] = str(total)
        ri[37] = '非課税'
        rows_out.append(ri)

        included.append({
            '患者名':    display_name,
            '合計':      total,
            '支払い方法': payment,
            '備考':      biko,
            '内訳':      breakdown,
        })

    # ── まとめて：1枚の請求書にまとめる ──────────────────────────
    if matome_patients:
        matome_patients.sort(key=lambda x: x[0])  # 五十音順（名前の文字コード順）
        total_all = sum(t for _, _, t in matome_patients)

        r = empty_row()
        r[0]  = '40101'
        r[1]  = '請求書'
        r[2]  = '夕陽ヶ丘苑'
        r[3]  = f'{billing_label} 薬局自己負担額'
        r[4]  = invoice_date
        r[5]  = payment_due
        r[7]  = billing_date
        r[10] = str(total_all)
        r[11] = '0'
        r[12] = str(total_all)
        r[13] = '御中'
        r[22] = ' '
        r[23] = ' '
        rows_out.append(r)

        for (pname, days, total) in matome_patients:
            breakdown = day_breakdown(days)
            ri = empty_row()
            ri[0]  = '40101'
            ri[1]  = '品目'
            ri[29] = f'{pname}　調剤報酬（自己負担）'
            ri[31] = str(total)
            ri[32] = '1'
            ri[35] = breakdown
            ri[36] = str(total)
            ri[37] = '非課税'
            rows_out.append(ri)

            included.append({
                '患者名':    pname,
                '合計':      total,
                '支払い方法': 'まとめて',
                '備考':      '夕陽ヶ丘苑 御中 にまとめて請求',
                '内訳':      breakdown,
            })

    buf = io.StringIO()
    csv.writer(buf).writerows(rows_out)
    return buf.getvalue().encode('utf-8-sig'), included, skipped


# ==================== サイドバー：設定 ====================
with st.sidebar:
    st.header('⚙️ 設定')

    today = date.today()
    default_year  = today.year  if today.month > 1 else today.year - 1
    default_month = today.month - 1 if today.month > 1 else 12

    billing_year      = st.number_input('請求対象年', value=default_year,
                                         min_value=2020, max_value=2100, step=1)
    billing_month_num = st.number_input('請求対象月', value=default_month,
                                         min_value=1, max_value=12, step=1)

    billing_label = f'令和{billing_year - 2018}年{billing_month_num}月分'
    last_day      = monthrange(billing_year, billing_month_num)[1]
    billing_date  = f'{billing_year}/{billing_month_num:02d}/{last_day:02d}'

    st.markdown(f'**対象月:** {billing_label}')
    st.markdown(f'**売上計上日:** {billing_date}')
    st.divider()

    invoice_date = st.date_input('請求日', value=today)
    payment_due  = st.date_input(
        'お支払期限',
        value=date(today.year, today.month, monthrange(today.year, today.month)[1]),
    )
    st.divider()

    sheets_url    = st.secrets.get('SHEETS_URL', '')
    sheet_gid     = st.secrets.get('SHEET_GID', '')
    furikomi_info = st.secrets.get('FURIKOMI_INFO', '')


# ==================== メインエリア ====================
col1, col2 = st.columns(2)

with col1:
    st.subheader('① 調剤報酬 PDF')
    med_file = st.file_uploader('患者別月間負担額一覧PDFをアップロード', type='pdf', key='med')

with col2:
    st.subheader('② 口座振替スケジュール PDF（ICC患者用）')

    DATA_DIR       = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    SAVED_SCHEDULE = os.path.join(DATA_DIR, '口座振替スケジュール.pdf')

    if billing_year == 2026:
        st.success('✅ 2026年スケジュール組み込み済み（アップロード不要）')
        schedule_file = None

    elif billing_year >= 2027:
        st.warning(f'⚠️ {billing_year}年のスケジュールPDFをアップロードしてください。')
        if os.path.exists(SAVED_SCHEDULE):
            mtime = date.fromtimestamp(os.path.getmtime(SAVED_SCHEDULE))
            st.info(f'保存済みスケジュールを使用します（更新日: {mtime}）')
            schedule_file = None
    
        else:
            schedule_file = st.file_uploader('年間スケジュールPDFをアップロード', type='pdf', key='sch')
    
        new_schedule = st.file_uploader(
            '📂 スケジュールを更新する',
            type='pdf', key='sch_update',
            help='新しいスケジュールPDFをアップロードすると data/ に上書き保存されます',
        )
        if new_schedule:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(SAVED_SCHEDULE, 'wb') as f:
                f.write(new_schedule.read())
            st.success('✅ スケジュールを更新しました！')
            st.rerun()
    else:
        if os.path.exists(SAVED_SCHEDULE):
            mtime = date.fromtimestamp(os.path.getmtime(SAVED_SCHEDULE))
            st.success(f'保存済みスケジュールを使用します（更新日: {mtime}）')
            schedule_file = None
    
        else:
            st.info('ICC（口座振替）患者がいる場合はアップロードしてください。')
            schedule_file = st.file_uploader('年間スケジュールPDFをアップロード', type='pdf', key='sch')
    
        new_schedule = st.file_uploader(
            '📂 スケジュールを更新する',
            type='pdf', key='sch_update',
            help='新しいスケジュールPDFをアップロードすると data/ に上書き保存されます',
        )
        if new_schedule:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(SAVED_SCHEDULE, 'wb') as f:
                f.write(new_schedule.read())
            st.success('✅ スケジュールを更新しました！')
            st.rerun()

st.divider()

# ==================== 実行ボタン ====================
if st.button('🚀 CSV を生成する', type='primary', use_container_width=True):
    if not med_file:
        st.error('調剤報酬PDFをアップロードしてください。')
    else:
        with st.spinner('処理中...'):

            # ── PDF パース ──────────────────────────────────────────
            try:
                pdf_data = parse_pdf_per_day(med_file.read())
            except Exception as e:
                st.error(f'PDF の読み込みエラー: {e}')
                st.stop()

            # ── スケジュール取得 ──────────────────────────────────
            transfer_text = None
            if billing_year == 2026:
                transfer_text = transfer_date_text(SCHEDULE_2026, billing_month_num)
            else:
                sch_source = SAVED_SCHEDULE if os.path.exists(SAVED_SCHEDULE) else None
                if schedule_file:
                    sch_source = schedule_file
                if sch_source:
                    try:
                        sch_bytes = (open(sch_source, 'rb').read()
                                     if isinstance(sch_source, str)
                                     else sch_source.read())
                        schedule      = parse_transfer_schedule(sch_bytes)
                        transfer_text = transfer_date_text(schedule, billing_month_num)
                    except Exception as e:
                        st.warning(f'スケジュールPDF の読み込みに失敗しました: {e}')
                elif billing_year >= 2027:
                    st.warning(f'⚠️ {billing_year}年のスケジュールPDFがないため、ICC患者の備考欄に振替日が記載されません。')

            # ── 会計シート取得 ───────────────────────────────────
            try:
                gid = sheet_gid.strip() if sheet_gid.strip() else None
                payment_data = fetch_kaike_sheet(sheets_url, gid=gid)
            except Exception as e:
                st.error(f'会計シートの取得エラー: {e}')
                st.stop()

            # ── CSV 生成 ─────────────────────────────────────────
            csv_bytes, included, skipped = build_csv(
                pdf_data, payment_data,
                billing_label,
                invoice_date.strftime('%Y/%m/%d'),
                payment_due.strftime('%Y/%m/%d'),
                billing_date,
                transfer_text,
                furikomi_info,
            )

        # ===== 結果表示 =====
        matome_count = sum(1 for x in included if x['支払い方法'] == 'まとめて')
        indiv_count  = len(included) - matome_count

        msg = f'✅ 完了！ 請求書対象 **{len(included)} 名**'
        if matome_count:
            msg += f'　（まとめて {matome_count} 名 → 夕陽ヶ丘苑 御中 1枚）'
        if transfer_text:
            msg += f'　口座振替日: **{transfer_text}**'
        st.success(msg)

        # ダウンロードボタン
        filename = f'請求書_{billing_year}{billing_month_num:02d}.csv'
        st.download_button(
            label=f'📥 {filename} をダウンロード',
            data=csv_bytes,
            file_name=filename,
            mime='text/csv',
            use_container_width=True,
        )

        st.divider()

        # 請求一覧テーブル
        if included:
            import pandas as pd
            st.subheader(f'📋 請求一覧（{len(included)} 名）')
            df = pd.DataFrame(included)
            df['合計'] = df['合計'].apply(lambda x: f'¥{x:,}')
            st.dataframe(df, use_container_width=True, hide_index=True)

        # 除外一覧（折りたたみ）
        if skipped:
            with st.expander(f'⚠️ 除外した患者（{len(skipped)} 名）'):
                import pandas as pd
                st.dataframe(pd.DataFrame(skipped), use_container_width=True,
                             hide_index=True)
