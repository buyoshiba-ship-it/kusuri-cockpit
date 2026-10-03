"""
update_mhlw_data.py
厚生労働省の公式ページから最新の「医療用医薬品供給状況」Excelを自動ダウンロードし、
Webアプリ（cockpit.html）用の data.js を自動更新＆差分履歴（status_history.json）を記録するスクリプト。
"""

import os
import re
import ssl
import json
import urllib.request
import openpyxl
import sys

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MHLW_PAGE_URL = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html"
DATA_JS_PATH = os.path.join(BASE_DIR, "data.js")
PREV_JSON_PATH = os.path.join(BASE_DIR, "prev_medicine_data.json")
DIFF_HISTORY_PATH = os.path.join(BASE_DIR, "diff_history.json")
DOWNLOAD_EXCEL_PATH = os.path.join(BASE_DIR, "mhlw_latest_supply.xlsx")

def get_ssl_context():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx

def fetch_latest_excel_url():
    print("🔍 厚生労働省の最新ページを確認中...")
    req = urllib.request.Request(
        MHLW_PAGE_URL, 
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    )
    ctx = get_ssl_context()
    
    with urllib.request.urlopen(req, context=ctx) as res:
        html = res.read().decode('utf-8', errors='ignore')
    
    matches = re.findall(r'<a[^>]+href=["\']([^"\']+\.xlsx)["\'][^>]*>(.*?)</a>', html, re.DOTALL | re.IGNORECASE)
    if not matches:
        raise Exception("厚労省ページからExcelファイルのリンクが見つかりませんでした。")
    
    for href, title in matches:
        clean_title = re.sub(r'<[^>]+>', '', title).strip()
        if "医療用医薬品供給状況" in clean_title or "kyoukyu" in href.lower() or "iyakuhin" in href.lower():
            full_url = urllib.parse.urljoin(MHLW_PAGE_URL, href)
            print(f"✅ 最新Excelを発見: {clean_title}")
            print(f"🔗 URL: {full_url}")
            return full_url, clean_title

    # マッチしなかった場合は最初のxlsxを採用
    href, title = matches[0]
    full_url = urllib.parse.urljoin(MHLW_PAGE_URL, href)
    print(f"✅ Excelを発見: {title.strip()} ({full_url})")
    return full_url, title.strip()

def download_excel(url):
    print(f"📥 最新Excelをダウンロード中...")
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    )
    ctx = get_ssl_context()
    with urllib.request.urlopen(req, context=ctx) as res, open(DOWNLOAD_EXCEL_PATH, 'wb') as f:
        f.write(res.read())
    size_mb = os.path.getsize(DOWNLOAD_EXCEL_PATH) / (1024 * 1024)
    print(f"✅ ダウンロード完了: {size_mb:.2f} MB")

def parse_excel_to_dataset(excel_path):
    print("⚙️ Excelデータを解析・構造化中...")
    wb = openpyxl.load_workbook(excel_path, read_only=True)
    ws = wb.active
    
    medicine_data = []
    
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i <= 1 or not row or len(row) < 14:
            continue
        
        name = str(row[5] or "").strip()
        ingredient = str(row[2] or "").strip()
        maker = str(row[6] or "").strip()
        yj_code = str(row[4] or "").strip()
        status_str = str(row[11] or "").strip()
        reason = str(row[13] or "").strip()
        recovery = str(row[14] or "").strip()
        category = str(row[9] or "").strip()
        
        # 更新日: col 12 (出荷対応更新日) または col 19 (その他更新日)
        update_date = ""
        if row[12] is not None and str(row[12]).strip() not in ["None", ""]:
            update_date = str(row[12]).strip()
        elif len(row) > 19 and row[19] is not None and str(row[19]).strip() not in ["None", ""]:
            update_date = str(row[19]).strip()
            
        ship_vol = str(row[16] or "").strip() if len(row) > 16 else ""
        
        # 製品区分・基礎的医薬品タグ
        raw_prod = str(row[7] or "")
        raw_basic = str(row[8] or "")
        
        tags = []
        if "先発品" in raw_prod:
            tags.append("先")
        elif "準先発品" in raw_prod:
            tags.append("準先")
        elif "後発品" in raw_prod:
            tags.append("後")
        elif "長期収載品" in raw_prod:
            tags.append("旧後")
            
        if "1：対象" in raw_basic or "対象" in raw_basic:
            tags.append("基")
            
        prod_tag = " ".join(tags)
        
        # 解除見込み時期 (col 15: 2027年1月など)
        rec_timing = str(row[15] or "").strip() if len(row) > 15 and row[15] is not None else ""
        if rec_timing in ["None", "－", "-"]:
            rec_timing = ""
            
        # 今回更新NEWマーク (col 20: 'New' の場合)
        is_new = 1 if len(row) > 20 and str(row[20] or "").strip() == "New" else 0
            
        # ステータス区分: 0=通常, 1=限定出荷, 2=供給停止
        st_type = 0
        if "供給停止" in status_str or "停止" in status_str:
            st_type = 2
        elif "限定出荷" in status_str:
            st_type = 1
        else:
            st_type = 0
            
        medicine_data.append([
            name,
            ingredient,
            maker,
            yj_code,
            status_str,
            reason,
            recovery,
            st_type,
            category,
            update_date,
            ship_vol,
            prod_tag,
            rec_timing,
            is_new
        ])
        
    print(f"✅ 全 {len(medicine_data):,} 品目の抽出に成功しました！")
    return medicine_data

def detect_diff_and_save(new_data):
    """前回データと比較して差分（改善・悪化・停止）を検知"""
    prev_map = {}
    if os.path.exists(PREV_JSON_PATH):
        try:
            with open(PREV_JSON_PATH, "r", encoding="utf-8") as f:
                prev_list = json.load(f)
                for item in prev_list:
                    yj = item[3]
                    if yj:
                        prev_map[yj] = item
        except Exception as e:
            print("⚠️ 前回データの読み込みスキップ:", e)

    diffs = {
        "improved": [],
        "worsened": [],
        "new_stopped": [],
        "total_new_items": 0
    }
    
    for row in new_data:
        yj = row[3]
        if not yj or yj not in prev_map:
            diffs["total_new_items"] += 1
            continue
            
        prev_row = prev_map[yj]
        prev_st = prev_row[7] # 0, 1, 2
        new_st = row[7]
        
        name = row[0]
        maker = row[2]
        
        # 改善（例: 限定出荷/停止 ➔ 通常、停止 ➔ 限定）
        if new_st < prev_st:
            diffs["improved"].append({
                "yj": yj,
                "name": name,
                "maker": maker,
                "from_status": prev_row[4],
                "to_status": row[4],
                "changeType": "improved"
            })
        # 悪化（通常 ➔ 限定/停止、限定 ➔ 停止）
        elif new_st > prev_st:
            is_stopped = (new_st == 2)
            diff_entry = {
                "yj": yj,
                "name": name,
                "maker": maker,
                "from_status": prev_row[4],
                "to_status": row[4],
                "changeType": "worsened"
            }
            diffs["worsened"].append(diff_entry)
            if is_stopped:
                diffs["new_stopped"].append(diff_entry)
                
    print(f"\n📊 差分検知結果:")
    print(f"  🟢 改善・通常復帰した品目: {len(diffs['improved'])} 件")
    print(f"  🔴 悪化・出荷制限になった品目: {len(diffs['worsened'])} 件 (うち供給停止: {len(diffs['new_stopped'])} 件)")
    
    # 差分履歴を保存
    with open(DIFF_HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(diffs, f, ensure_ascii=False, indent=2)
        
    # 今回のデータを次回比較用として保存
    with open(PREV_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(new_data, f, ensure_ascii=False)

def export_data_js(medicine_data):
    print("💾 data.js を出力中...")
    json_str = json.dumps(medicine_data, ensure_ascii=False)
    js_content = f"const MEDICINE_DATA = {json_str};\n"
    
    with open(DATA_JS_PATH, "w", encoding="utf-8") as f:
        f.write(js_content)
        
    size_mb = os.path.getsize(DATA_JS_PATH) / (1024 * 1024)
    print(f"🎉 data.js の更新が完了しました！（ファイルサイズ: {size_mb:.2f} MB）")

def main():
    print("=" * 60)
    print(" 🏥 厚生労働省 医薬品供給データ 自動同期パイプライン")
    print("=" * 60)
    
    try:
        excel_url, title = fetch_latest_excel_url()
        download_excel(excel_url)
        dataset = parse_excel_to_dataset(DOWNLOAD_EXCEL_PATH)
        detect_diff_and_save(dataset)
        export_data_js(dataset)
        print("\n✨ すべての工程が正常に完了しました！")
    except Exception as e:
        print(f"\n❌ エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
