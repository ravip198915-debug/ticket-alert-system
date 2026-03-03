import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import io

st.set_page_config(page_title="Telegram Ticket Dashboard", layout="wide")

# ==============================
# CUSTOM CSS (Sticky Header + Styling)
# ==============================
st.markdown("""
<style>
.header {
    position: sticky;
    top: 0;
    background-color: white;
    padding: 15px;
    z-index: 999;
    border-bottom: 2px solid #D32F2F;
}
.metric-card {
    background-color: white;
    padding: 15px;
    border-radius: 10px;
    box-shadow: 0px 2px 6px rgba(0,0,0,0.1);
}
.stButton>button {
    background-color: #D32F2F;
    color: white;
    font-weight: bold;
    border-radius: 8px;
    height: 3em;
    width: 100%;
}
</style>
""", unsafe_allow_html=True)

# ==============================
# STICKY HEADER
# ==============================
st.markdown("""
<div class="header">
    <h2>📢 Telegram Ticket Alert Dashboard</h2>
</div>
""", unsafe_allow_html=True)

st.write("")

# ==============================
# LAYOUT
# ==============================
col1, col2 = st.columns([1, 2])

with col1:
    st.subheader("🔐 Configuration")

    BOT_TOKEN = st.text_input("Enter Telegram Bot Token", type="password")

    ticket_file = st.file_uploader("Upload Ticket System Excel", type=["xlsx"])
    master_file = st.file_uploader("Upload TE_MASTER Excel", type=["xlsx"])

    send_button = st.button("🚀 Send Alerts")

with col2:
    st.subheader("📊 Live Summary")

# ==============================
# PROCESS FILES
# ==============================
if ticket_file and master_file:

    tickets = pd.read_excel(ticket_file)
    engineers = pd.read_excel(master_file)

    total_engineers = len(engineers)
    total_tickets = len(tickets)
    total_p1 = tickets[tickets["Priority"].astype(str).str.upper() == "P1"].shape[0]

    m1, m2, m3 = st.columns(3)

    with m1:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("👨‍🔧 Engineers", total_engineers)
        st.markdown('</div>', unsafe_allow_html=True)

    with m2:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("🎫 Total Tickets", total_tickets)
        st.markdown('</div>', unsafe_allow_html=True)

    with m3:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.metric("🔴 P1 Tickets", total_p1)
        st.markdown('</div>', unsafe_allow_html=True)

    st.divider()

    st.subheader("📍 Taluk Wise Ticket Count")
    summary = tickets.groupby("Taluk").size().reset_index(name="Tickets")
    st.dataframe(summary, use_container_width=True)

    st.bar_chart(summary.set_index("Taluk"))

# ==============================
# TELEGRAM FUNCTION
# ==============================
def send_telegram(chat_id, message):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    response = requests.post(url, data={
        "chat_id": int(chat_id),
        "text": message,
        "parse_mode": "HTML"
    })
    return response

# ==============================
# SEND ALERTS
# ==============================
if send_button and ticket_file and master_file:

    if not BOT_TOKEN:
        st.error("Enter Bot Token")
        st.stop()

    success = 0
    failed = 0
    sent_log = []

    progress = st.progress(0)

    for i, eng_row in engineers.iterrows():

        taluk = eng_row["Taluk"]
        chat_id = eng_row["Telegram_ID"]
        eng_name = eng_row.get("Engineer Name", "")

        if pd.isna(chat_id):
            continue

        taluk_tickets = tickets[tickets["Taluk"] == taluk]

        if taluk_tickets.empty:
            continue

        msg = f"🕒 {datetime.now().strftime('%d %B %Y | %I:%M %p')}\n\n"
        msg += f"🚨 OPEN TICKETS\n\n👨‍🔧 {eng_name}\n📍 {taluk}\n\n"

        for _, row in taluk_tickets.iterrows():
            msg += f"🆔 {row.get('Ticket Number','')}\n"
            msg += f"⚠ {row.get('Priority','')}\n"
            msg += "-----------------\n"

        response = send_telegram(chat_id, msg)

        if response.status_code == 200:
            success += 1
            status = "SUCCESS"
        else:
            failed += 1
            status = "FAILED"

        sent_log.append({
            "Date_Time": datetime.now(),
            "Engineer": eng_name,
            "Taluk": taluk,
            "Status": status
        })

        progress.progress((i+1)/len(engineers))

    st.divider()

    s1, s2 = st.columns(2)
    s1.metric("✅ Success", success)
    s2.metric("❌ Failed", failed)

    log_df = pd.DataFrame(sent_log)
    st.dataframe(log_df, use_container_width=True)

    output = io.BytesIO()
    log_df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)

    st.download_button(
        label="📥 Download Log File",
        data=output,
        file_name=f"Sent_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    st.success("Process Completed Successfully ✅")