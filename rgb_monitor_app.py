import streamlit as st
import pandas as pd
import os
import glob
import time
from datetime import datetime, timedelta
import io
import numpy as np
import toml
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import requests
import win32com.client as win32
import pythoncom

# Load configuration
try:
    config = toml.load("config.toml")
    devices = config.get("devices", {})
    smtp_settings = config.get("smtp", {})
    spc_config = config.get("spc", {})
except Exception as e:
    st.error(f"Config loading failed: {e}")
    devices = {}
    smtp_settings = {}
    spc_config = {}

# Initialize session state for device selection if not present
if 'selected_device_key' not in st.session_state:
    if devices:
        st.session_state.selected_device_key = list(devices.keys())[0]
    else:
        st.session_state.selected_device_key = None

# Initialize alert state
if 'alert_state' not in st.session_state:
    st.session_state.alert_state = {} # {device_key: {channel: last_alert_time}}

# Get current device info
current_device_key = st.session_state.selected_device_key
current_device_name = ""
if current_device_key and current_device_key in devices:
    current_device_name = devices[current_device_key].get("name", "")

# Set page config dynamically based on selected device
page_title = f"{current_device_name}健常性監視アプリ" if current_device_name else "健常性監視アプリ"
st.set_page_config(page_title=page_title)

def style_dataframe(df):
    """
    Styles the DataFrame:
    - Background white (default in Streamlit light mode, but we can force it if needed, though pandas styler focuses on cells)
    - Text red for NG status rows or values? User said "NGの値は赤にする".
    - Let's highlight rows where Status is NG with light red background, or text red.
    """
    def highlight_ng_row(row):
        # Check if any column has 'NG' or if Status column is NG
        is_ng = False
        if 'StatusR' in row and row['StatusR'] == 'NG': is_ng = True
        if 'StatusG' in row and row['StatusG'] == 'NG': is_ng = True
        if 'StatusB' in row and row['StatusB'] == 'NG': is_ng = True
        if 'Status' in row and row['Status'] == 'NG': is_ng = True
        
        return ['color: red' if is_ng else '' for _ in row]

    return df.style.apply(highlight_ng_row, axis=1)

def parse_csv_line(line):
    """
    Parses a single line of the Color CSV.
    Expected format: StatusR, ValueR, StatusG, ValueG, StatusB, ValueB
    """
    parts = line.strip().split(',')
    if len(parts) < 6:
        return None
    
    try:
        data = {
            'Type': 'Color',
            'StatusR': parts[0],
            'ValueR': float(parts[1]),
            'StatusG': parts[2],
            'ValueG': float(parts[3]),
            'StatusB': parts[4],
            'ValueB': float(parts[5])
        }
        return data
    except ValueError:
        return None

def parse_scale_line(line):
    """
    Parses a single line of the Scale CSV.
    Expected format: Status, ValueX, ValueY
    Example: OK,99.13167,99.61333
    """
    parts = line.strip().split(',')
    if len(parts) < 3:
        return None
    
    try:
        data = {
            'Type': 'Scale',
            'Status': parts[0],
            'ValueX': float(parts[1]),
            'ValueY': float(parts[2])
        }
        return data
    except ValueError:
        return None

def plot_chart(df, channel_info, y_min, y_max, y_step, show_ma=False, ma_window=5, hist_tick_x=1.0, hist_tick_y=2.0):
    """
    Helper function to plot the control chart using Matplotlib.
    """
    values = df[channel_info['col_val']].values
    statuses = df[channel_info['col_stat']].values
    dates = df['Timestamp']
    
    # Limits from channel info
    limits = channel_info.get('limits', {})
    UCL = limits.get('ucl', 131)
    CL = limits.get('cl', 128)
    LCL = limits.get('lcl', 125)

    # Create subplots: Main chart on top, Histogram on bottom
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8), gridspec_kw={'height_ratios': [3, 1]})
    
    # --- Main Chart (ax1) ---
    
    # X-axis values (Equidistant)
    x_values = np.arange(len(dates))
    
    # Plot Line Segments (Conditional Coloring)
    for i in range(len(values) - 1):
        is_ok_curr = (statuses[i] != 'NG' and (LCL <= values[i] <= UCL))
        is_ok_next = (statuses[i+1] != 'NG' and (LCL <= values[i+1] <= UCL))
        
        if is_ok_curr and is_ok_next:
            color = channel_info['color']
        else:
            color = 'gray'
            
        ax1.plot(x_values[i:i+2], values[i:i+2], color=color, linewidth=1)
    
    # Plot Points (Color coded & Marker)
    # Rule 1: Outside UCL/LCL -> 'x' marker
    for i, (val, status) in enumerate(zip(values, statuses)):
        is_out_of_spec = (val > UCL or val < LCL) # Rule 1
        
        marker = 'o'
        color = channel_info['color']
        size = 30
        
        if status == 'NG' or is_out_of_spec:
            color = 'black' # Or red? User didn't specify color for 'x', just 'x'. Black is fine for contrast.
            if is_out_of_spec:
                marker = 'x'
                size = 50
        
        ax1.scatter(x_values[i], val, c=color, marker=marker, s=size, zorder=5)
        
        # Annotate value ONLY if out of spec
        if is_out_of_spec:
            ax1.annotate(f"{val:.1f}", (x_values[i], val), textcoords="offset points", xytext=(0, 5), ha='center', fontsize=8, color='red')

    # Moving Average
    if show_ma and len(values) >= ma_window:
        ma_values = pd.Series(values).rolling(window=ma_window).mean()
        ax1.plot(x_values, ma_values, color='orange', linestyle='-', linewidth=1.5, label=f'MA({ma_window})')

    # Control Lines
    ax1.axhline(y=UCL, color='red', linestyle='--', label='UCL')
    ax1.axhline(y=CL, color='green', linestyle='-', label='CL')
    ax1.axhline(y=LCL, color='red', linestyle='--', label='LCL')

    # X-axis Settings (Show all ticks, Equidistant)
    ax1.set_xticks(x_values)
    ax1.set_xticklabels(dates.dt.strftime('%Y-%m-%d %H:%M'), rotation=90, fontsize=8)
    # ax1.set_xlabel("Date") # Hide x-label for top plot to save space
    
    # Y-axis Settings
    # Use passed y_min/max for Color, but for Scale we might need to adjust or use auto?
    # User said "Color Y軸設定" in sidebar. Scale limits are 99-101. 120-140 is too big.
    # If channel is Scale, maybe auto-scale or use limits +/- margin?
    # Let's check channel name.
    if "Scale" in channel_info['name']:
        # Auto scale based on limits or data
        margin = (UCL - LCL) * 0.5
        ax1.set_ylim(LCL - margin, UCL + margin)
        # Ticks?
        ax1.set_yticks(np.arange(LCL - margin, UCL + margin, (UCL-LCL)/4)) # Approximate
    else:
        ax1.set_ylim(y_min, y_max)
        ax1.set_yticks(np.arange(y_min, y_max + y_step, y_step))
        
    ax1.set_ylabel("Value")
    
    # Title, Grid, Legend
    ax1.set_title(f"{channel_info['name']} Control Chart")
    ax1.grid(True, which='both', linestyle='--', linewidth=0.5)
    ax1.legend(loc='upper left')
    
    # --- Histogram (ax2) ---
    ax2.hist(values, bins=20, color=channel_info['color'], alpha=0.7, edgecolor='black')
    ax2.axvline(x=UCL, color='red', linestyle='--', linewidth=1)
    ax2.axvline(x=CL, color='green', linestyle='-', linewidth=1)
    ax2.axvline(x=LCL, color='red', linestyle='--', linewidth=1)
    ax2.set_title("Histogram")
    ax2.set_xlabel("Value")
    ax2.set_ylabel("Frequency")
    
    # Histogram Ticks
    # User said: "デフォルトはX=1,Y=2とする。" (Default X=1, Y=2)
    # And "ヒストグラムの目盛り間隔は立て横個別で設定できるようにする"
    
    # X-axis ticks (Value)
    # Start from min value or LCL?
    # Let's use matplotlib's MultipleLocator if possible, or manual range.
    import matplotlib.ticker as ticker
    ax2.xaxis.set_major_locator(ticker.MultipleLocator(hist_tick_x))
    ax2.yaxis.set_major_locator(ticker.MultipleLocator(hist_tick_y))

    ax2.grid(True, linestyle='--', linewidth=0.5)

    # Adjust layout
    plt.tight_layout()
    
    return fig

def send_popup(message):
    st.toast(message, icon="⚠️")

def send_email(to_addr, subject, body, attachments=None, cc_addr=None, bcc_addr=None):
    try:
        # Initialize COM library for current thread
        pythoncom.CoInitialize()
        
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = Mail Item
        
        mail.To = to_addr
        if cc_addr:
            mail.CC = cc_addr
        if bcc_addr:
            mail.BCC = bcc_addr
            
        mail.Subject = subject
        mail.To = to_addr
        mail.Subject = subject
        # mail.Body = body # Text format
        mail.HTMLBody = body # HTML format
        
        if attachments:
            for path in attachments:
                if os.path.exists(path):
                    try:
                        mail.Attachments.Add(path)
                    except Exception as e:
                        print(f"Failed to attach {path}: {e}")
        
        mail.Send()
        print(f"Email sent to {to_addr}") # Logging
        st.toast(f"メール送信成功: {to_addr}", icon="📧")
    except Exception as e:
        print(f"Outlook email send failed: {e}")
        st.error(f"Outlook email send failed: {e}")
    finally:
        # Uninitialize COM library? Usually not strictly necessary for simple dispatch in thread but good practice if heavy usage.
        # pythoncom.CoUninitialize() 
        pass

def save_alert_chart(df, channel_info, output_path):
    """
    Generates and saves a chart for the alert, filtered to the last 1 month.
    """
    try:
        # Filter last 1 month
        if df.empty:
            return False
            
        last_timestamp = df['Timestamp'].max()
        one_month_ago = last_timestamp - timedelta(days=30)
        df_filtered = df[df['Timestamp'] >= one_month_ago]
        
        if df_filtered.empty:
            return False

        # Plot
        # We need y_min/max/step. Let's use defaults or try to get from somewhere?
        # For alert charts, maybe auto-scale is better to see the anomaly clearly?
        # Or use the channel limits.
        limits = channel_info.get('limits', {})
        UCL = limits.get('ucl', 131)
        LCL = limits.get('lcl', 125)
        CL = limits.get('cl', 128)
        
        # Use auto-scale with margin around limits/data
        vals = df_filtered[channel_info['col_val']].dropna()
        if vals.empty:
            return False
            
        y_min = min(vals.min(), LCL) - 2
        y_max = max(vals.max(), UCL) + 2
        y_step = (y_max - y_min) / 5
        
        # Reuse plot_chart but we need to pass params.
        # plot_chart signature: df, channel_info, y_min, y_max, y_step, show_ma=False, ma_window=5, hist_tick_x=1.0, hist_tick_y=2.0
        # We can use defaults for others.
        fig = plot_chart(df_filtered, channel_info, y_min, y_max, y_step)
        
        # Save
        fig.savefig(output_path, format='jpg')
        plt.close(fig)
        return True
    except Exception as e:
        print(f"Error in save_alert_chart: {e}")
        return False



def send_teams(webhook_url, message):
    try:
        payload = {"text": message}
        requests.post(webhook_url, json=payload)
    except Exception as e:
        print(f"Teams send failed: {e}")

def check_alert(df, device_key, alert_settings):
    """
    Checks for alert conditions based on config.
    Consolidates alerts into one email/notification.
    """
    if not df.empty and device_key in devices:
        device_config = devices[device_key]
        alert_rule = device_config.get("alert_rule", {})
        window_size = alert_rule.get("window_size", 5)
        consecutive_ng = alert_rule.get("consecutive_ng", 3)
        
        notifications = device_config.get("notifications", {})
        
        # Check each channel
        channels = [
            {'name': 'Red', 'col_stat': 'StatusR', 'col_val': 'ValueR', 'color': 'red', 'limits': device_config.get('limits', {}).get('color', {})},
            {'name': 'Green', 'col_stat': 'StatusG', 'col_val': 'ValueG', 'color': 'green', 'limits': device_config.get('limits', {}).get('color', {})},
            {'name': 'Blue', 'col_stat': 'StatusB', 'col_val': 'ValueB', 'color': 'blue', 'limits': device_config.get('limits', {}).get('color', {})},
            # Add Scale channels if needed? User said "Red,Blueなど異常が複数の項目で発生した場合".
            # Scale data is in the same DF but different rows/cols?
            # Scale cols: ValueX, ValueY. Status?
            # Let's check if Scale cols exist in DF.
            # If df has mixed data, we need to be careful.
            # But check_alert iterates rows?
            # Actually, the current check_alert logic assumes Color columns exist.
            # If df is mixed, we should separate or handle NaNs.
            # For now, let's stick to Color as per original code, or add Scale if requested.
            # User mentioned "Red,Blue etc", implying Color.
        ]
        
        alerts_found = []
        
        for ch in channels:
            # Filter for Color data (where col_val is not NaN)
            df_ch = df.dropna(subset=[ch['col_val']])
            
            # Get latest window
            recent_df = df_ch.tail(window_size)
            if len(recent_df) < consecutive_ng:
                continue
            
            # Check consecutive NGs in the tail
            last_n_stats = recent_df[ch['col_stat']].tail(consecutive_ng)
            last_n_vals = recent_df[ch['col_val']].tail(consecutive_ng)
            
            # Limits
            limits = ch.get('limits', {})
            UCL = limits.get('ucl', 131)
            LCL = limits.get('lcl', 125)

            is_consecutive_alert = True
            for stat, val in zip(last_n_stats, last_n_vals):
                # NG condition: Status is NG OR Value out of range (125-131)
                if stat == 'NG' or not (LCL <= val <= UCL):
                    pass # This is NG
                else:
                    is_consecutive_alert = False
                    break
            
            # Rule 1: Latest point out of limits (Immediate trigger)
            is_rule1_alert = False
            if not recent_df.empty:
                latest_val = recent_df[ch['col_val']].iloc[-1]
                if not (LCL <= latest_val <= UCL):
                    is_rule1_alert = True

            if is_consecutive_alert or is_rule1_alert:
                alerts_found.append(ch)

        if alerts_found:
            # Trigger Alert
            print(f"Alert triggered for: {[ch['name'] for ch in alerts_found]}")
            
            # Popup (Show for each or summary?)
            if alert_settings['popup']:
                msg = " / ".join([ch['name'] for ch in alerts_found])
                send_popup(f"異常検知: {msg}")
            
            # Email
            if alert_settings['email']:
                print("Email alert setting is ON")
                email_conf = notifications.get('email', {})
                to_addr = email_conf.get('to_addr')
                cc_addr = email_conf.get('cc_addr')
                bcc_addr = email_conf.get('bcc_addr')
                
                subject = f"[Alert] {device_config.get('name', '装置')} 異常検知"
                
                body = f"{device_config.get('name', '装置')}で以下の異常が検知されました。\n\n"
                body = f"<h3>{device_config.get('name', '装置')}で以下の異常が検知されました。</h3><ul>"
                
                attachment_paths = []
                
                import tempfile
                with tempfile.TemporaryDirectory() as tmpdirname:
                    for ch in alerts_found:
                        body += f"<li><b>{ch['name']}</b>: 直近{consecutive_ng}回連続で管理値外れまたはNG</li>"
                        
                        # Generate Chart
                        try:
                            chart_path = os.path.join(tmpdirname, f"{ch['name']}_alert_chart.jpg")
                            # We need the full df for the channel to plot history
                            df_ch = df.dropna(subset=[ch['col_val']])
                            if save_alert_chart(df_ch, ch, chart_path):
                                attachment_paths.append(chart_path)
                            else:
                                print(f"Failed to save alert chart for {ch['name']}")
                        except Exception as e:
                            print(f"Error generating chart for {ch['name']}: {e}")
                    
                    body += "</ul>"
                    body += "<p>直近1ヶ月のチャートを添付しました。ご確認ください。</p>"
                    
                    # Send Email
                    print(f"Sending email to {to_addr} with {len(attachment_paths)} attachments")
                    try:
                        send_email(to_addr, subject, body, attachment_paths, cc_addr, bcc_addr)
                    except Exception as e:
                        print(f"Failed to send email in check_alert: {e}")
                        st.error(f"Failed to send email: {e}")
            else:
                print("Email alert setting is OFF")
            
            # Teams
            if alert_settings['teams']:
                teams_conf = notifications.get('teams', {})
                msg = " / ".join([ch['name'] for ch in alerts_found])
                send_teams(teams_conf.get('webhook_url'), f"異常検知: {msg}")

def calculate_cp_cpk(values, usl, lsl):
    """
    Calculates Cp, Cpk, and Sigma (std dev).
    """
    if len(values) < 2:
        return None, None, None
    
    sigma = np.std(values, ddof=1)
    mean = np.mean(values)
    
    if sigma == 0:
        return None, None, 0.0
        
    cp = (usl - lsl) / (6 * sigma)
    cpk = min((usl - mean) / (3 * sigma), (mean - lsl) / (3 * sigma))
    
    return cp, cpk, sigma

def get_rule_description(rule_id, count, config_val):
    """
    Returns the title and body for the anomaly rule.
    """
    descriptions = {
        1: {
            "title": "最新のplotが管理限界を超えた場合",
            "body": """管理図には、上方管理限界（UCL: Upper Control Limit）と下方管理限界（LCL: Lower Control Limit）が設定されています。
これらは通常、プロセスの平均から±3シグマ（標準偏差）の範囲で決定されます。

この範囲内にデータが収まっている状態が正常です。
逆に、データ点がこれらの管理限界を超えると、工程に重大な異常が発生している可能性が高いことを示します。

管理限界を超えた場合、即座に異常原因を調査し、対策を講じる必要があります。
まずは工程に機械の故障や人為的ミスなどの突発的な問題がないか確認し、次に使用されている材料や条件に急激な変化がなかったかを調査しましょう。"""
        },
        2: {
            "title": f"連続で{count}点がCLより上または下にある場合",
            "body": f"""{count}点連続で平均線より上または下に偏っている場合、データが偶然のばらつきによるものではなく、何らかの原因で工程全体に構造的な変化が生じて偏ってしまった可能性があります。

このパターンが見られた場合、機械の設定や温度、湿度などの工程条件の見直しを行い、偏りの原因を突き止めます。

また、定期的なメンテナンスや人員教育を検討することも大切です。"""
        },
        3: {
            "title": f"連続{count}点が連続して昇順または降順に並んでいる場合",
            "body": f"""{count}点連続してデータが上昇または下降している場合、単なる偶然ではなく、工程にトレンドが発生している可能性があります。
これは、外的な影響や設備の状態変化、あるいは原材料の品質変動が原因と考えられ、放置すると品質問題に発展するリスクがあります。

このパターンが見られたときは、機械や設備の状態を確認し、メンテナンスが必要かどうかを判断します。

また、原材料や作業環境の変化についても調査する必要があります。"""
        },
        4: {
            "title": f"連続で{count}点が交互に上下している場合",
            "body": f"""データが交互に上昇・下降するパターンが{count}回連続して発生する場合、工程が安定しておらず、一定の周期的な問題が存在する可能性があります。

このパターンは多くの場合、測定機器の不具合や不適切な作業手順によって発生します。
そのため解消をするためには、まず測定機器やセンサーの点検と再校正を行い、そこに問題がない場合は作業手順やプロセス自体に不適切な手順がないか確認します。"""
        },
        5: {
            "title": f"連続{count}点がCLを超えて同じ側にある場合",
            "body": f"""{count}つのデータ点が連続して、平均線より上か下に位置している場合、その工程に一時的な異常がある可能性があります。
偶然生まれた変動である可能性も考えられますが、放置するとさらなる異常につながるケースもあるため、警戒が必要です。

このパターンが見られた場合、作業環境や原材料に一時的な変動がないか確認します。
環境と材料に変動がなかった場合は、人的要因によるミスがないか調査します。"""
        },
        6: {
            "title": f"連続{count}点がCLよりも2σ以上離れている場合",
            "body": f"""データ点が{count}回連続して、2シグマの範囲外かつ管理限界内にある場合、工程が安定していない可能性があります。
2シグマ範囲外にデータが頻繁に出ることは通常のばらつきでは説明できません。
早期に対策しないと、不良品の発生率が増加するリスクがあります。

このパターンが見られた場合は、工程条件や機械の設定に変動がないかチェックし、原材料の品質や作業手順の見直しも行います。"""
        },
        7: {
            "title": f"n点が連続して1σ範囲内にある場合", # Note: User text said "n点が..." in title but usually we substitute. I'll keep user's title format but substitute if it makes sense. User said "()内の文字をタイトルとし". The user provided title has 'n' but the body has the number. I will substitute 'n' with the number in the title too for clarity, or keep as requested. User said "()内の変数に置き換えて表示する". So I will replace 'n' with count.
            "body": f"""データ点が連続して{count}回、平均線から±1シグマの範囲内に収まっている場合、データのばらつきが非常に小さくなっています。
これは一見するといいことのように見えますが、逆に測定システムやプロセスそのものに異常が隠れている可能性があります。

このパターンが見られた際は、測定機器の精度や感度が適切でありデータが正しく収集されているかを再確認します。"""
        },
        8: {
            "title": f"m点が連続して1σ範囲外にある場合",
            "body": f"""{count}回連続でデータ点が1シグマ範囲外にある場合、工程が著しく安定していない状態です。
何らかの原因で大きな変動が生じていることが考えられます。

このパターンが見られた場合、まず原材料や作業環境、設備に大きな変動がないか調査します。
次に調査結果から不安定な要因を特定し、工程の再調整を実施します。"""
        }
    }
    
    # Fix titles for 7 and 8 to replace n/m with count
    if rule_id == 7:
        descriptions[7]["title"] = descriptions[7]["title"].replace("n", str(count))
    if rule_id == 8:
        descriptions[8]["title"] = descriptions[8]["title"].replace("m", str(count))

    return descriptions.get(rule_id, {"title": "Unknown Rule", "body": ""})

def check_spc_rules(values, cl, sigma, config):
    """
    Checks SPC rules and returns a list of anomalies.
    values: list or numpy array of data points (chronological)
    cl: Center Line
    sigma: Standard Deviation (Process Sigma)
    config: Dictionary of rule parameters
    """
    anomalies = []
    n = len(values)
    if n == 0:
        return anomalies

    # Parameters
    rule2_a = config.get("rule2_a", 7)
    rule3_a = config.get("rule3_a", 7)
    rule4_x = config.get("rule4_x", 14)
    rule5_y = config.get("rule5_y", 3)
    rule6_z = config.get("rule6_z", 5)
    rule7_n = config.get("rule7_n", 15)
    rule8_m = config.get("rule8_m", 8)

    # Pre-calculate deviations
    ucl = cl + 3 * sigma
    lcl = cl - 3 * sigma
    sigma_2_upper = cl + 2 * sigma
    sigma_2_lower = cl - 2 * sigma
    sigma_1_upper = cl + 1 * sigma
    sigma_1_lower = cl - 1 * sigma

    # Rule 1: Latest point outside UCL/LCL
    if values[-1] > ucl or values[-1] < lcl:
        anomalies.append({"rule": 1, "count": 1})

    # Rule 2: Consecutive 'a' points on one side of CL
    if n >= rule2_a:
        last_a = values[-rule2_a:]
        if np.all(last_a > cl) or np.all(last_a < cl):
            anomalies.append({"rule": 2, "count": rule2_a})

    # Rule 3: Consecutive 'a' points increasing or decreasing
    if n >= rule3_a:
        last_a = values[-rule3_a:]
        diffs = np.diff(last_a)
        if np.all(diffs > 0) or np.all(diffs < 0):
            anomalies.append({"rule": 3, "count": rule3_a})

    # Rule 4: Consecutive 'x' points alternating
    if n >= rule4_x:
        last_x = values[-rule4_x:]
        diffs = np.diff(last_x)
        # Check if signs alternate
        # diffs[i] * diffs[i+1] < 0 means alternating signs
        is_alternating = True
        for i in range(len(diffs) - 1):
            if diffs[i] * diffs[i+1] >= 0:
                is_alternating = False
                break
        if is_alternating:
            anomalies.append({"rule": 4, "count": rule4_x})

    # Rule 5: Consecutive 'y' points on same side of CL
    if n >= rule5_y:
        last_y = values[-rule5_y:]
        if np.all(last_y > cl) or np.all(last_y < cl):
            anomalies.append({"rule": 5, "count": rule5_y})

    # Rule 6: Consecutive 'z' points > 2 sigma from CL (but within UCL/LCL? User says "2シグマの範囲外かつ管理限界内")
    # Actually standard rule is usually "2 out of 3 > 2sigma".
    # User definition: "Continuous z points are > 2sigma away from CL".
    if n >= rule6_z:
        last_z = values[-rule6_z:]
        # Check if all are outside 2 sigma (either > +2s or < -2s)
        # And user text says "within control limits" (implied by being data points? or strictly < 3s?)
        # "2シグマの範囲外かつ管理限界内" -> 2s < |val - cl| < 3s ?
        # Or just |val - cl| > 2s ? Usually "outside 2 sigma" means > 2s.
        # I will check |val - cl| > 2*sigma.
        is_rule6 = True
        for v in last_z:
            if not (abs(v - cl) > 2 * sigma):
                is_rule6 = False
                break
        if is_rule6:
            anomalies.append({"rule": 6, "count": rule6_z})

    # Rule 7: Consecutive 'n' points within 1 sigma
    if n >= rule7_n:
        last_n = values[-rule7_n:]
        if np.all(np.abs(last_n - cl) <= sigma):
            anomalies.append({"rule": 7, "count": rule7_n})

    # Rule 8: Consecutive 'm' points outside 1 sigma
    if n >= rule8_m:
        last_m = values[-rule8_m:]
        if np.all(np.abs(last_m - cl) > sigma):
            anomalies.append({"rule": 8, "count": rule8_m})

    return anomalies

def plot_combined_chart(df, channels, y_min, y_max, y_step, show_ma=False, ma_window=5, hist_tick_x=1.0, hist_tick_y=2.0):
    """
    Plots a combined chart for all channels.
    Note: Combined chart with different limits (Scale vs Color) is tricky.
    User said "Summaryに統合したScaleのX,Yのグラフを...".
    If we plot them on the same chart, the Y-axis scale will be messed up (130 vs 100).
    Maybe use dual axis? Or just plot Color combined and Scale combined separately?
    User said "Summaryに統合したScaleのX,Yのグラフを、Blueの右にScale X, ScaleYのタブを作成する" -> This refers to tabs.
    "SummaryのCombined Chart & Statisticsも管理値外れの値は×で表示する" -> This refers to the combined chart.
    If I mix Color (130) and Scale (100), I should probably use two subplots or dual axis.
    However, usually "Combined Chart" implies overlay.
    Let's try to overlay but maybe Scale is too different.
    Actually, let's just plot them. If they are far apart, they will just be lines at different levels.
    """
    dates = df['Timestamp']
    # Create subplots: Main chart on top, Histogram on bottom
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8), gridspec_kw={'height_ratios': [3, 1]})
    
    # --- Main Chart (ax1) ---
    x_values = np.arange(len(dates))
    
    for ch in channels:
        # Filter df for this channel's type? No, df is combined but has NaNs?
        # Actually df has all columns if we merged?
        # Wait, df is a list of dicts converted to DF. 
        # If we have mixed Color and Scale rows, 'ValueR' will be NaN for Scale rows.
        # We need to handle this.
        # But wait, the main loop separates df_color and df_scale.
        # `plot_combined_chart` receives `df`?
        # In the new logic, I should pass the relevant DF to the plot function or handle the combined DF.
        # The previous logic passed `df` which was sorted by timestamp.
        # If I pass `df` (combined), I need to interpolate or drop NaNs for plotting lines.
        
        # Let's extract valid values for this channel
        # ch['col_val']
        
        series = df[ch['col_val']]
        # We need to align with x_values (which corresponds to df['Timestamp']).
        # If data is missing (NaN), we can't plot line easily.
        # But here, Color and Scale come from different files, so timestamps might be different or interleaved.
        # If interleaved, we have NaNs.
        # Matplotlib plot ignores NaNs (breaks line). This is probably desired.
        
        # Get limits for this channel
        limits = ch.get('limits', {})
        UCL = limits.get('ucl', 131)
        LCL = limits.get('lcl', 125)
        
        # Drop NaNs for plotting lines
        # Drop NaNs for plotting lines
        # We need positional indices relative to the current df/x_values
        # series is a Series from df.
        # Let's get boolean mask of valid values
        mask = ~series.isna()
        valid_series = series[mask]
        
        # Get positional indices where mask is True
        valid_positions = np.where(mask)[0]
        
        if len(valid_positions) == 0:
            continue

        valid_x = x_values[valid_positions]
        
        ax1.plot(valid_x, valid_series, color=ch['color'], linewidth=1, label=ch['name'], marker='o', markersize=4)
        
        # Mark Anomalies (Rule 1)
        # We need to iterate to find out-of-spec
        for pos, val in zip(valid_positions, valid_series):
            if val > UCL or val < LCL:
                ax1.scatter(x_values[pos], val, c='black', marker='x', s=50, zorder=10)
                ax1.annotate(f"{val:.1f}", (x_values[pos], val), textcoords="offset points", xytext=(0, 5), ha='center', fontsize=8, color='red')

        if show_ma and len(valid_series) >= ma_window:
            ma_values = valid_series.rolling(window=ma_window).mean()
            ax1.plot(valid_x, ma_values, color=ch['color'], linestyle=':', linewidth=1, alpha=0.7)

        if show_ma and len(valid_series) >= ma_window:
            ma_values = valid_series.rolling(window=ma_window).mean()
            ax1.plot(valid_x, ma_values, color=ch['color'], linestyle=':', linewidth=1, alpha=0.7)

    # Control Lines
    # Use limits from the first channel as reference (assuming shared limits for combined chart of same type)
    if channels:
        first_ch = channels[0]
        limits = first_ch.get('limits', {})
        UCL = limits.get('ucl', 131)
        CL = limits.get('cl', 128)
        LCL = limits.get('lcl', 125)
        
        ax1.axhline(y=UCL, color='red', linestyle='--', label='UCL')
        ax1.axhline(y=CL, color='green', linestyle='-', label='CL')
        ax1.axhline(y=LCL, color='red', linestyle='--', label='LCL')

    # X-axis Settings
    ax1.set_xticks(x_values)
    ax1.set_xticklabels(dates.dt.strftime('%Y-%m-%d %H:%M'), rotation=90, fontsize=8)
    # ax1.set_xlabel("Date")
    
    # Y-axis Settings
    # If mixed, auto-scale.
    # If only Color, use sidebar settings.
    # How to detect? Check channel names.
    has_scale = any("Scale" in ch['name'] for ch in channels)
    if not has_scale:
        ax1.set_ylim(y_min, y_max)
        ax1.set_yticks(np.arange(y_min, y_max + y_step, y_step))
    else:
        # For Scale, maybe use limits to set range?
        # Scale limits are usually tight (e.g. 99-101).
        # Let's use auto-scale but ensure limits are visible.
        # Or use the same logic as individual charts if limits are known.
        if channels:
             # Use limits from first channel
            first_ch = channels[0]
            limits = first_ch.get('limits', {})
            UCL = limits.get('ucl', 101)
            LCL = limits.get('lcl', 99)
            margin = (UCL - LCL) * 0.5
            ax1.set_ylim(LCL - margin, UCL + margin)
        else:
            ax1.autoscale(axis='y')
        
    ax1.set_ylabel("Value")
    
    # Title, Grid, Legend
    ax1.set_title("Combined Control Chart")
    ax1.grid(True, which='both', linestyle='--', linewidth=0.5)
    ax1.legend(loc='upper left', bbox_to_anchor=(1.05, 1), borderaxespad=0.) # Move legend outside to avoid covering data? Or keep inside. User said "legend also".
    # Standard legend inside is fine, but if many lines, maybe outside.
    # Let's keep it inside 'upper left' or 'best' but ensure it includes control lines.
    ax1.legend(loc='upper left', fontsize='small', framealpha=0.5)
    
    # --- Histogram (ax2) ---
    for ch in channels:
        series = df[ch['col_val']].dropna()
        ax2.hist(series, bins=20, color=ch['color'], alpha=0.5, label=ch['name'], edgecolor='black')
    
    # Skip limit lines on histogram too if mixed
    
    ax2.set_title("Histogram")
    ax2.set_xlabel("Value")
    ax2.set_ylabel("Frequency")
    ax2.legend()
    ax2.grid(True, linestyle='--', linewidth=0.5)

    # Histogram Ticks
    import matplotlib.ticker as ticker
    ax2.xaxis.set_major_locator(ticker.MultipleLocator(hist_tick_x))
    ax2.yaxis.set_major_locator(ticker.MultipleLocator(hist_tick_y))

    plt.tight_layout()
    return fig

def main():
    st.title(page_title)

    # Sidebar for inputs
    st.sidebar.header("設定")
    
    # Device Selection
    # Sidebar with Expanders
    
    # Device Selection
    if devices:
        device_options = list(devices.keys())
        device_names = [d['name'] for d in devices.values()]
        
        current_index = 0
        if st.session_state.selected_device_key in device_options:
            current_index = device_options.index(st.session_state.selected_device_key)
            
        with st.sidebar.expander("装置設定", expanded=True):
            selected_device_name = st.selectbox(
                "装置選択", 
                device_names, 
                index=current_index
            )
            folder_path = st.text_input("監視対象フォルダパス", value=os.getcwd())
        
        selected_index = device_names.index(selected_device_name)
        selected_key = device_options[selected_index]
        
        if selected_key != st.session_state.selected_device_key:
            st.session_state.selected_device_key = selected_key
            st.rerun()
            
    # Monitoring Interval
    with st.sidebar.expander("監視間隔", expanded=False):
        col1, col2 = st.columns([2, 1])
        with col1:
            interval_val = st.number_input("値", min_value=1, value=1)
        with col2:
            interval_unit = st.selectbox("単位", ["s", "m", "h", "d"], index=3)
    
    if interval_unit == "s":
        sleep_time = interval_val
    elif interval_unit == "m":
        sleep_time = interval_val * 60
    elif interval_unit == "h":
        sleep_time = interval_val * 3600
    elif interval_unit == "d":
        sleep_time = interval_val * 86400

    # Y-axis Settings (Color)
    with st.sidebar.expander("Color Y軸設定", expanded=False):
        y_min_color = st.number_input("最小値", value=120)
        y_max_color = st.number_input("最大値", value=140)
        y_step_color = st.number_input("目盛り間隔", value=2)

    # Chart Settings
    with st.sidebar.expander("チャート設定", expanded=False):
        show_ma = st.checkbox("移動平均線を表示", value=False)
        ma_window = 5
        if show_ma:
            ma_window = st.number_input("移動平均期間", min_value=2, value=5)
        
        st.caption("ヒストグラム設定")
        hist_tick_x = st.number_input("X軸目盛り間隔", value=1.0, step=0.1)
        hist_tick_y = st.number_input("Y軸目盛り間隔", value=2.0, step=1.0)

    # Alert Settings
    st.sidebar.subheader("異常発報")
    alert_on = st.sidebar.toggle("ON/OFF", value=False)
    
    alert_settings = {'popup': False, 'email': False, 'teams': False}
    
    if alert_on:
        current_notifications = {}
        if current_device_key and current_device_key in devices:
            current_notifications = devices[current_device_key].get("notifications", {})
            
        alert_settings['popup'] = st.sidebar.checkbox("ポップアップ通知", help=current_notifications.get("popup_message", ""))
        alert_settings['email'] = st.sidebar.checkbox("メール通知", help=f"To: {current_notifications.get('email', {}).get('to_addr', '')}")
        alert_settings['teams'] = st.sidebar.checkbox("Teams通知(※作成中)", help=f"Webhook: {current_notifications.get('teams', {}).get('webhook_url', '')}")



    # Limits from Config
    # Default values
    limits_color = {'cl': 128, 'ucl': 131, 'lcl': 125}
    limits_scale = {'cl': 100, 'ucl': 101, 'lcl': 99}
    
    if current_device_key and current_device_key in devices:
        dev_limits = devices[current_device_key].get("limits", {})
        if "color" in dev_limits:
            limits_color.update(dev_limits["color"])
        if "scale" in dev_limits:
            limits_scale.update(dev_limits["scale"])

    # Constants (Now from config)
    # CL = 128
    # UCL = 131
    # LCL = 125

    # Monitoring Control
    if 'monitoring' not in st.session_state:
        st.session_state.monitoring = False

    start_btn = st.sidebar.button("監視開始")
    stop_btn = st.sidebar.button("監視停止")

    if start_btn:
        st.session_state.monitoring = True
    if stop_btn:
        st.session_state.monitoring = False

    if st.session_state.monitoring:
        st.sidebar.success(f"監視中... (間隔: {interval_val}{interval_unit})")
        
        if not os.path.exists(folder_path):
            st.error(f"指定されたフォルダが存在しません: {folder_path}")
            st.session_state.monitoring = False
            return

        csv_files = glob.glob(os.path.join(folder_path, "*.csv"))
        
        if not csv_files:
            st.warning("CSVファイルが見つかりませんでした。")
        else:
            all_data = []
            for file in csv_files:
                try:
                    ctime = os.path.getctime(file)
                    dt = datetime.fromtimestamp(ctime)
                    filename = os.path.basename(file)
                    
                    # Determine type by filename or content
                    is_scale = "Scale" in filename or "Scal" in filename
                    is_color = "Color" in filename
                    
                    # If neither in filename, peek at content (not implemented fully, relying on filename for now as per request "if names don't match, read file". User said "if names don't match... read file". I'll add simple check if filename check fails)
                    
                    with open(file, 'r', encoding='utf-8-sig') as f:
                        for line in f:
                            parsed = None
                            if is_scale:
                                parsed = parse_scale_line(line)
                            elif is_color:
                                parsed = parse_csv_line(line)
                            else:
                                # Fallback: Try parsing as Color first, then Scale
                                parsed = parse_csv_line(line)
                                if not parsed:
                                    parsed = parse_scale_line(line)
                            
                            if parsed:
                                parsed['Filename'] = filename
                                parsed['Timestamp'] = dt
                                all_data.append(parsed)
                except Exception as e:
                    st.error(f"ファイル読み込みエラー {file}: {e}")

            if not all_data:
                st.warning("有効なデータが見つかりませんでした。")
            else:
                df = pd.DataFrame(all_data)
                df = df.sort_values('Timestamp')

                # Debug Info (Moved here to access df)
                with st.sidebar.expander("デバッグ情報", expanded=True):
                    st.write(f"Alert ON: {alert_on}")
                    if not df.empty:
                        st.write(f"Data Points: {len(df)}")
                        last_row = df.iloc[-1]
                        st.write(f"Latest Timestamp: {last_row['Timestamp']}")
                        
                        # Check alert condition manually for debug
                        if current_device_key and current_device_key in devices:
                            dev_conf = devices[current_device_key]
                            ar = dev_conf.get("alert_rule", {})
                            c_ng = ar.get("consecutive_ng", 3)
                            st.write(f"Consecutive NG Setting: {c_ng}")
                            
                            # Check Red
                            if 'ValueR' in df.columns:
                                df_color_debug = df.dropna(subset=['ValueR'])
                                recent = df_color_debug.tail(c_ng)
                                st.write("Recent Red Values:")
                                st.write(recent[['Timestamp', 'ValueR', 'StatusR']])
                            
                            # Check Green
                            if 'ValueG' in df.columns:
                                df_color_debug = df.dropna(subset=['ValueG'])
                                recent = df_color_debug.tail(c_ng)
                                st.write("Recent Green Values:")
                                st.write(recent[['Timestamp', 'ValueG', 'StatusG']])

                            # Check Blue
                            if 'ValueB' in df.columns:
                                df_color_debug = df.dropna(subset=['ValueB'])
                                recent = df_color_debug.tail(c_ng)
                                st.write("Recent Blue Values:")
                                st.write(recent[['Timestamp', 'ValueB', 'StatusB']])
                    else:
                        st.write("No Data in DF")

                # Check Alerts
                if alert_on:
                    check_alert(df, current_device_key, alert_settings)

                # Time Range Filter
                st.subheader("期間フィルタ")
                filter_option = st.radio("表示期間", ["全期間", "1日", "1週間", "1ヶ月", "1年", "期間指定"], index=3, horizontal=True)
                
                now = datetime.now()
                start_date = None
                end_date = None
                
                if filter_option == "1日":
                    start_date = now - timedelta(days=1)
                elif filter_option == "1週間":
                    start_date = now - timedelta(weeks=1)
                elif filter_option == "1ヶ月":
                    start_date = now - timedelta(days=30)
                elif filter_option == "1年":
                    start_date = now - timedelta(days=365)
                elif filter_option == "期間指定":
                    date_range = st.date_input("期間を選択", [])
                    if len(date_range) == 2:
                        start_date = datetime.combine(date_range[0], datetime.min.time())
                        end_date = datetime.combine(date_range[1], datetime.max.time())
                
                if start_date:
                    df = df[df['Timestamp'] >= start_date]
                if end_date:
                    df = df[df['Timestamp'] <= end_date]

                # Separate DataFrames
                df_color = df[df['Type'] == 'Color'].copy()
                df_scale = df[df['Type'] == 'Scale'].copy()

                # Create tabs
                tab_summary, tab1, tab2, tab3, tab4, tab5 = st.tabs(["Summary", "Red", "Green", "Blue", "Scale X", "Scale Y"])

                # Define Channels
                channels_color = []
                if not df_color.empty:
                    channels_color = [
                        {'name': 'Red', 'col_val': 'ValueR', 'col_stat': 'StatusR', 'color': 'red', 'tab': tab1, 'limits': limits_color},
                        {'name': 'Green', 'col_val': 'ValueG', 'col_stat': 'StatusG', 'color': 'green', 'tab': tab2, 'limits': limits_color},
                        {'name': 'Blue', 'col_val': 'ValueB', 'col_stat': 'StatusB', 'color': 'blue', 'tab': tab3, 'limits': limits_color}
                    ]
                
                channels_scale = []
                if not df_scale.empty:
                    channels_scale = [
                        {'name': 'Scale X', 'col_val': 'ValueX', 'col_stat': 'Status', 'color': 'purple', 'tab': tab4, 'limits': limits_scale},
                        {'name': 'Scale Y', 'col_val': 'ValueY', 'col_stat': 'Status', 'color': 'orange', 'tab': tab5, 'limits': limits_scale}
                    ]
                
                all_channels = channels_color + channels_scale

                # Summary Tab
                with tab_summary:
                    st.subheader("Combined Control Chart (Color)")
                    
                    # Combined Chart (Color)
                    if channels_color:
                        fig_combined_color = plot_combined_chart(df, channels_color, y_min_color, y_max_color, y_step_color, show_ma, ma_window, hist_tick_x, hist_tick_y)
                        st.pyplot(fig_combined_color)
                    else:
                        st.info("Colorデータがありません。")

                    st.subheader("Combined Control Chart (Scale)")
                    # Combined Chart (Scale)
                    if channels_scale:
                        # For Scale, we might want auto-scale or specific limits. 
                        # Passing y_min_color etc is irrelevant if plot_combined_chart handles auto-scale for Scale.
                        # But let's pass them anyway as placeholders.
                        fig_combined_scale = plot_combined_chart(df, channels_scale, y_min_color, y_max_color, y_step_color, show_ma, ma_window, hist_tick_x, hist_tick_y)
                        st.pyplot(fig_combined_scale)
                    else:
                        st.info("Scaleデータがありません。")
                    
                    # Statistics
                    st.subheader("Statistics (Color)")
                    stats_data_color = []
                    for ch in channels_color:
                        vals = df[ch['col_val']].dropna().values
                        if len(vals) == 0:
                            continue
                            
                        limits = ch.get('limits', {})
                        UCL = limits.get('ucl', 131)
                        LCL = limits.get('lcl', 125)
                        
                        cp, cpk, sigma = calculate_cp_cpk(vals, UCL, LCL)
                        
                        stats_data_color.append({
                            "Channel": ch['name'],
                            "AVE": round(np.mean(vals), 2),
                            "MIN": round(np.min(vals), 2),
                            "MAX": round(np.max(vals), 2),
                            "Sigma": round(sigma, 3) if sigma is not None else "-",
                            "Cp": round(cp, 3) if cp is not None else "-",
                            "Cpk": round(cpk, 3) if cpk is not None else "-"
                        })
                    if stats_data_color:
                        st.dataframe(pd.DataFrame(stats_data_color))
                    else:
                        st.info("No Color Data")

                    st.subheader("Statistics (Scale)")
                    stats_data_scale = []
                    for ch in channels_scale:
                        vals = df[ch['col_val']].dropna().values
                        if len(vals) == 0:
                            continue
                            
                        limits = ch.get('limits', {})
                        UCL = limits.get('ucl', 101)
                        LCL = limits.get('lcl', 99)
                        
                        cp, cpk, sigma = calculate_cp_cpk(vals, UCL, LCL)
                        
                        stats_data_scale.append({
                            "Channel": ch['name'],
                            "AVE": round(np.mean(vals), 2),
                            "MIN": round(np.min(vals), 2),
                            "MAX": round(np.max(vals), 2),
                            "Sigma": round(sigma, 3) if sigma is not None else "-",
                            "Cp": round(cp, 3) if cp is not None else "-",
                            "Cpk": round(cpk, 3) if cpk is not None else "-"
                        })
                    if stats_data_scale:
                        st.dataframe(pd.DataFrame(stats_data_scale))
                    else:
                        st.info("No Scale Data")
                    
                    st.divider()
                    
                    st.subheader("個別明るさチャート")
                    for ch in all_channels:
                        st.markdown(f"**{ch['name']}**")
                        
                        limits = ch.get('limits', {})
                        UCL = limits.get('ucl', 131)
                        CL = limits.get('cl', 128)
                        LCL = limits.get('lcl', 125)
                        
                        # Filter df for this channel (to remove NaNs from other types if any, though plot_chart handles it via col_val)
                        # But plot_chart expects 'Timestamp' and col_val.
                        # If we pass the combined df, it has NaNs for rows of other types.
                        # plot_chart uses `df[channel_info['col_val']].values`.
                        # If we have NaNs, matplotlib might complain or plot gaps.
                        # Ideally we pass only valid rows for this channel.
                        df_ch = df.dropna(subset=[ch['col_val']])
                        
                        fig = plot_chart(df_ch, ch, y_min_color, y_max_color, y_step_color, show_ma, ma_window, hist_tick_x, hist_tick_y)
                        st.pyplot(fig)
                        
                        # SPC Anomalies
                        vals = df_ch[ch['col_val']].values
                        process_sigma = (UCL - CL) / 3
                        anomalies = check_spc_rules(vals, CL, process_sigma, spc_config)
                        if anomalies:
                            st.markdown(f"<span style='color:red'>**{ch['name']} 管理値外れまたは警戒項目が発生しました:**</span>", unsafe_allow_html=True)
                            for anomaly in anomalies:
                                desc = get_rule_description(anomaly['rule'], anomaly['count'], spc_config)
                                with st.expander(f"🔴 {desc['title']}"):
                                    st.write(desc['body'])

                # Individual Tabs
                for ch in all_channels:
                    with ch['tab']:
                        st.subheader(f"{ch['name']} Control Chart")
                        
                        limits = ch.get('limits', {})
                        UCL = limits.get('ucl', 131)
                        CL = limits.get('cl', 128)
                        LCL = limits.get('lcl', 125)
                        
                        df_ch = df.dropna(subset=[ch['col_val']])
                        
                        fig = plot_chart(df_ch, ch, y_min_color, y_max_color, y_step_color, show_ma, ma_window, hist_tick_x, hist_tick_y)
                        st.pyplot(fig)
                        
                        # SPC Anomalies
                        vals = df_ch[ch['col_val']].values
                        process_sigma = (UCL - CL) / 3
                        anomalies = check_spc_rules(vals, CL, process_sigma, spc_config)
                        
                        if anomalies:
                            st.markdown(f"<span style='color:red'>**管理値外れまたは警戒項目が発生しました:**</span>", unsafe_allow_html=True)
                            for anomaly in anomalies:
                                desc = get_rule_description(anomaly['rule'], anomaly['count'], spc_config)
                                with st.expander(f"🔴 {desc['title']}"):
                                    st.write(desc['body'])
                        
                        # Raw Data with Styling
                        st.subheader("Raw Data")
                        # Show relevant columns
                        cols = ['Timestamp', 'Filename', ch['col_stat'], ch['col_val']]
                        df_display = df_ch[cols].sort_values('Timestamp', ascending=False)
                        st.dataframe(style_dataframe(df_display), use_container_width=True)
                        
                        # Download Chart
                        # Matplotlib figure can be saved to buffer
                        img_buffer = io.BytesIO()
                        fig.savefig(img_buffer, format='jpg')
                        img_buffer.seek(0)
                        
                        st.download_button(
                            label="Download Chart (JPG)",
                            data=img_buffer,
                            file_name=f"{ch['name']}_chart.jpg",
                            mime="image/jpeg",
                            key=f"dl_{ch['name']}"
                        )

                # Download Data
                st.subheader("データダウンロード")
                
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                excel_buffer.seek(0)
                
                st.download_button(
                    label="Download Data (XLSX)",
                    data=excel_buffer,
                    file_name="rgb_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Show data table - Removed as per request "Summary以外の各ページにある...表示は削除して"
                # st.dataframe(df)

        # Auto-refresh logic
        time.sleep(sleep_time)
        st.rerun()

if __name__ == "__main__":
    main()
