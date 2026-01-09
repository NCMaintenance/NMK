import streamlit as st
import streamlit.components.v1 as components
import pypff
import openpyxl
from openpyxl.utils.exceptions import IllegalCharacterError
import hashlib
import tempfile
import os
import pandas as pd
from datetime import datetime
import json
import collections

# --- Configuration ---
st.set_page_config(
    page_title="HSE Email Archive Extractor",
    page_icon="https://www.hse.ie/favicon-32x32.png",
    layout="wide"
)

# --- HTML TEMPLATE (HSE DASHBOARD STYLE) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
        
        :root {
            --hse-teal: #02594C;
            --hse-teal-light: #037362;
            --glass-bg: rgba(255, 255, 255, 0.9);
            --glass-border: rgba(255, 255, 255, 0.5);
            --primary-gradient: linear-gradient(135deg, #02594C 0%, #014D42 100%);
        }
        
        body {
            font-family: 'Inter', sans-serif;
            background-color: transparent; /* Transparent so it blends with Streamlit if needed */
        }
        
        .pro-card {
            background: var(--glass-bg);
            backdrop-filter: blur(20px);
            border: 1px solid var(--glass-border);
            transition: all 0.3s ease;
            border-radius: 1rem;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        }
        
        .pro-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 15px -3px rgba(2, 89, 76, 0.1);
            border-color: rgba(2, 89, 76, 0.2);
        }

        .header-gradient {
            background: var(--primary-gradient);
            box-shadow: 0 10px 40px -10px rgba(2, 89, 76, 0.4);
            border-bottom: 4px solid var(--hse-teal-light);
        }

        .stat-value {
            background: -webkit-linear-gradient(45deg, #02594C, #047857);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
    </style>
</head>
<body class="p-2">
    
    <!-- HEADER -->
    <header class="mb-6 px-6 py-5 rounded-2xl shadow-xl flex flex-col sm:flex-row justify-between items-center header-gradient text-white">
        <div class="flex items-center mb-4 sm:mb-0">
            <div class="bg-white/20 backdrop-blur-md p-3 rounded-xl mr-4">
                <!-- HSE Logo SVG Placeholder or Image -->
                <img src="https://www.hse.ie/image-library/hse-site-logo-2021.svg" alt="HSE Logo" class="h-12">
            </div>
            <div>
                <h1 class="text-2xl font-black tracking-tight">Email Archive Analytics</h1>
                <p class="text-teal-100 text-sm font-medium">Digital Forensics & Data Extraction</p>
            </div>
        </div>
        <div class="text-right">
            <div class="text-xs font-bold uppercase tracking-wider opacity-80">Report Generated</div>
            <div class="text-lg font-bold" id="current-date"></div>
        </div>
    </header>

    <!-- SCORECARDS -->
    <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
        <!-- Total Emails -->
        <div class="pro-card p-6 border-l-4 border-teal-600">
            <div class="flex justify-between items-start">
                <div>
                    <div class="text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">Total Emails</div>
                    <div class="text-5xl font-black text-gray-800" id="total-count">0</div>
                    <div class="text-xs text-teal-600 font-bold mt-2 bg-teal-50 inline-block px-2 py-1 rounded">Processed Successfully</div>
                </div>
                <div class="bg-teal-100 p-3 rounded-xl">
                    <svg class="w-6 h-6 text-teal-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"></path></svg>
                </div>
            </div>
        </div>

        <!-- Duplicates -->
        <div class="pro-card p-6 border-l-4 border-amber-500">
            <div class="flex justify-between items-start">
                <div>
                    <div class="text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">Duplicates Removed</div>
                    <div class="text-5xl font-black text-gray-800" id="duplicate-count">0</div>
                    <div class="text-xs text-amber-600 font-bold mt-2 bg-amber-50 inline-block px-2 py-1 rounded">Redundant Data</div>
                </div>
                <div class="bg-amber-100 p-3 rounded-xl">
                    <svg class="w-6 h-6 text-amber-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 7v8a2 2 0 002 2h6M8 7V5a2 2 0 012-2h4.586a1 1 0 01.707.293l4.414 4.414a1 1 0 01.293.707V15a2 2 0 01-2 2h-2M8 7H6a2 2 0 00-2 2v10a2 2 0 002 2h8a2 2 0 002-2v-2"></path></svg>
                </div>
            </div>
        </div>

        <!-- Folders -->
        <div class="pro-card p-6 border-l-4 border-blue-500">
            <div class="flex justify-between items-start">
                <div>
                    <div class="text-xs font-bold text-gray-500 uppercase tracking-widest mb-1">Folders Scanned</div>
                    <div class="text-5xl font-black text-gray-800" id="folder-count">0</div>
                    <div class="text-xs text-blue-600 font-bold mt-2 bg-blue-50 inline-block px-2 py-1 rounded">Structure Depth</div>
                </div>
                <div class="bg-blue-100 p-3 rounded-xl">
                    <svg class="w-6 h-6 text-blue-700" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z"></path></svg>
                </div>
            </div>
        </div>
    </div>

    <!-- CHARTS GRID -->
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
        <!-- Timeline Chart -->
        <div class="pro-card p-6">
            <h3 class="text-lg font-bold text-gray-800 mb-4 flex items-center">
                <div class="w-1.5 h-6 bg-teal-600 mr-3 rounded-full"></div>
                Email Volume Over Time
            </h3>
            <div class="h-64 relative">
                <canvas id="timelineChart"></canvas>
            </div>
        </div>

        <!-- Folder Distribution -->
        <div class="pro-card p-6">
            <h3 class="text-lg font-bold text-gray-800 mb-4 flex items-center">
                <div class="w-1.5 h-6 bg-blue-600 mr-3 rounded-full"></div>
                Top Folders by Volume
            </h3>
            <div class="h-64 relative">
                <canvas id="folderChart"></canvas>
            </div>
        </div>
    </div>

    <!-- FOOTER -->
    <footer class="text-center text-xs text-gray-400 font-medium py-4 border-t border-gray-200">
        <p>HSE Digital Services | Generated by Secure PST Extractor</p>
    </footer>

    <script>
        // --- DATA INJECTION ---
        const DATA = %%DATA_PLACEHOLDER%%;

        // --- INIT ---
        document.getElementById('current-date').textContent = new Date().toLocaleDateString('en-IE', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
        
        // --- ANIMATE NUMBERS ---
        function animateValue(id, start, end, duration) {
            if (start === end) return;
            const range = end - start;
            let current = start;
            const increment = end > start ? 1 : -1;
            const stepTime = Math.abs(Math.floor(duration / range));
            const obj = document.getElementById(id);
            const timer = setInterval(function() {
                current += increment;
                obj.innerHTML = current;
                if (current == end) { clearInterval(timer); }
            }, Math.max(stepTime, 10)); // Min 10ms for performance
        }

        animateValue("total-count", 0, DATA.stats.total, 1500);
        animateValue("duplicate-count", 0, DATA.stats.duplicates, 1500);
        animateValue("folder-count", 0, DATA.stats.folders, 1000);

        // --- CHART: TIMELINE ---
        const ctxTimeline = document.getElementById('timelineChart').getContext('2d');
        new Chart(ctxTimeline, {
            type: 'line',
            data: {
                labels: DATA.charts.timeline.labels,
                datasets: [{
                    label: 'Emails',
                    data: DATA.charts.timeline.values,
                    borderColor: '#02594C',
                    backgroundColor: 'rgba(2, 89, 76, 0.1)',
                    borderWidth: 2,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 3,
                    pointBackgroundColor: '#fff',
                    pointBorderColor: '#02594C'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: 'rgba(0,0,0,0.8)',
                        titleFont: { family: 'Inter', size: 13 },
                        bodyFont: { family: 'Inter', size: 13 },
                        padding: 10,
                        displayColors: false
                    }
                },
                scales: {
                    y: { grid: { borderDash: [2, 4], color: '#f0f0f0' }, beginAtZero: true },
                    x: { grid: { display: false } }
                }
            }
        });

        // --- CHART: FOLDERS ---
        const ctxFolder = document.getElementById('folderChart').getContext('2d');
        new Chart(ctxFolder, {
            type: 'bar',
            data: {
                labels: DATA.charts.folders.labels,
                datasets: [{
                    label: 'Email Count',
                    data: DATA.charts.folders.values,
                    backgroundColor: [
                        '#02594C', '#037362', '#047857', '#059669', '#10b981', '#34d399'
                    ],
                    borderRadius: 6
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: 'rgba(0,0,0,0.8)',
                        padding: 10
                    }
                },
                scales: {
                    x: { grid: { borderDash: [2, 4] } },
                    y: { grid: { display: false } }
                }
            }
        });
    </script>
</body>
</html>
"""

# --- AUTHENTICATION ---
def check_password():
    """Returns `True` if the user had the correct password."""
    # Look for password in secrets, otherwise default to a simple one for first-run ease
    correct_password = st.secrets.get("APP_PASSWORD", "hseadmin") 

    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    # Login Form
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        with st.container():
            st.image("https://www.hse.ie/image-library/hse-site-logo-2021.svg", width=150)
            st.markdown("### Secure PST Extractor")
            st.info("Please authenticate to access this application.")
            
            password = st.text_input("Password", type="password")
            if st.button("Log In", type="primary"):
                if password == correct_password:
                    st.session_state.password_correct = True
                    st.rerun()
                else:
                    st.error("😕 Incorrect password")
    return False

# --- HELPER FUNCTIONS ---

def clean_string(text):
    if not text: return ""
    return "".join(c for c in str(text) if c in ['\n', '\r', '\t'] or c >= ' ')

def format_date_time(timestamp):
    if not timestamp: return None, None
    try:
        if isinstance(timestamp, datetime):
            return timestamp.strftime("%d/%m/%Y"), timestamp.strftime("%H:%M:%S"), timestamp
        dt_str = str(timestamp)
        dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S.%f")
        return dt.strftime("%d/%m/%Y"), dt.strftime("%H:%M:%S"), dt
    except:
        return str(timestamp), "", None

def get_recipients(message):
    to_l, cc_l, bcc_l = [], [], []
    try:
        for i in range(message.get_number_of_recipients()):
            r = message.get_recipient(i)
            n, e = r.get_name(), r.get_email_address()
            entry = f"{n} <{e}>" if n and e else (e or n or "")
            if not entry: continue
            
            t = r.get_type()
            if t == 1: to_l.append(entry)
            elif t == 2: cc_l.append(entry)
            elif t == 3: bcc_l.append(entry)
    except: pass
    return "; ".join(to_l), "; ".join(cc_l), "; ".join(bcc_l)

def process_folder(folder, rows_list, current_path, seen_emails, progress_bar, status_text, stats_folders):
    folder_name = folder.get_name() or "Root"
    full_path = f"{current_path} > {folder_name}" if current_path else folder_name
    
    # Update Stats
    stats_folders[folder_name] = stats_folders.get(folder_name, 0)
    
    status_text.text(f"Scanning: {full_path}...")

    for message in folder.sub_messages:
        try:
            stats_folders[folder_name] += 1
            subject = clean_string(message.get_subject())
            body = clean_string(message.get_plain_text_body())
            
            d_obj = message.get_delivery_time() or message.get_client_submit_time()
            d_str, t_str, dt_obj = format_date_time(d_obj)
            
            to_s, cc_s, bcc_s = get_recipients(message)
            
            # Signature for deduping
            sig = hashlib.md5(f"{subject}|{d_str}|{t_str}|{body}".encode('utf-8', errors='ignore')).hexdigest()
            is_dup = "Yes" if sig in seen_emails else "No"
            if is_dup == "No": seen_emails.add(sig)

            rows_list.append({
                'Folder Path': full_path,
                'Folder Name': folder_name, # Helper for charts
                'Subject': subject,
                'Date': d_str,
                'Time': t_str,
                'To': clean_string(to_s),
                'CC': clean_string(cc_s),
                'BCC': clean_string(bcc_s),
                'Is Duplicate': is_dup,
                'Body': body,
                'DateTimeObj': dt_obj # Hidden column for sorting/charts
            })
        except: continue

    for sub in folder.sub_folders:
        process_folder(sub, rows_list, full_path, seen_emails, progress_bar, status_text, stats_folders)

# --- MAIN APP ---

def main():
    if not check_password():
        return

    # Sidebar
    with st.sidebar:
        st.image("https://www.hse.ie/image-library/hse-site-logo-2021.svg", width=100)
        st.markdown("### Control Panel")
        uploaded_file = st.file_uploader("Upload PST File", type=["pst"])
        st.info("Files are processed locally in memory.")
        st.markdown("---")
        st.markdown("**User:** Admin")
        st.markdown("**Unit:** Digital Services")

    st.title("Secure PST Extractor")
    st.markdown("Welcome to the HSE Secure Email Extractor. Upload a PST file to begin analysis.")

    if uploaded_file is not None:
        if st.button("Start Extraction", type="primary"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pst") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name

            try:
                # Init
                seen = set()
                rows = []
                folder_stats = {}
                progress = st.progress(0)
                status = st.empty()

                # Process
                pst = pypff.file()
                pst.open(tmp_path)
                process_folder(pst.get_root_folder(), rows, "", seen, progress, status, folder_stats)
                pst.close()
                
                progress.progress(100)
                status.empty()

                # Data Processing
                df = pd.DataFrame(rows)
                
                # --- PREPARE DASHBOARD DATA ---
                total_emails = len(df)
                duplicates = len(df[df['Is Duplicate'] == "Yes"])
                
                # Timeline Data (Group by Month)
                df['DateTimeObj'] = pd.to_datetime(df['DateTimeObj'], errors='coerce')
                timeline_data = df.dropna(subset=['DateTimeObj']).set_index('DateTimeObj').resample('M').size()
                timeline_labels = [d.strftime('%b %Y') for d in timeline_data.index]
                timeline_values = timeline_data.values.tolist()

                # Folder Data (Top 5)
                top_folders = dict(collections.Counter(folder_stats).most_common(6))
                # Remove "Root" if empty or useless
                if "Root" in top_folders and top_folders["Root"] == 0: del top_folders["Root"]
                
                dashboard_data = {
                    "stats": {
                        "total": total_emails,
                        "duplicates": duplicates,
                        "folders": len(folder_stats)
                    },
                    "charts": {
                        "timeline": {"labels": timeline_labels, "values": timeline_values},
                        "folders": {"labels": list(top_folders.keys()), "values": list(top_folders.values())}
                    }
                }

                # --- RENDER DASHBOARD ---
                # Inject JSON data into HTML
                html_code = HTML_TEMPLATE.replace("%%DATA_PLACEHOLDER%%", json.dumps(dashboard_data))
                components.html(html_code, height=850, scrolling=True)

                # --- DOWNLOAD SECTION ---
                st.markdown("### Export Data")
                col1, col2 = st.columns(2)
                
                # Clean DF for export (remove helper columns)
                export_df = df.drop(columns=['DateTimeObj', 'Folder Name'])
                
                from io import BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    export_df.to_excel(writer, index=False)
                
                with col1:
                    st.download_button(
                        label="📥 Download Excel Report",
                        data=output.getvalue(),
                        file_name=f"HSE_Extraction_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
            except Exception as e:
                st.error(f"Error: {e}")
            finally:
                if os.path.exists(tmp_path): os.unlink(tmp_path)

if __name__ == "__main__":
    main()
