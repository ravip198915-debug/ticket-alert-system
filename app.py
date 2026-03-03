import streamlit as st
import pandas as pd
import requests
from datetime import datetime
import os

# ==============================
# PAGE CONFIG
# ==============================
st.set_page_config(page_title="Telegram Ticket Alert System", layout="wide")

st.title("📢 Telegram Ticket Alert System")

# ==============================
# BOT TOKEN INPUT
# ==============================
BOT_TOKEN = st.text_input("Enter Telegram Bot Token", type="password")

# ==============================
# FILE UPLOAD
# ==============================
ticket_file = st.file_uploader("Upload Ticket System Excel", type=["xlsx"])
master_file = st.file_uploader("Upload TE_MASTER Excel", type=["xlsx"])

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
# MAIN BUTTON
# ==============================
if st.button("🚀 Send Alerts"):

    if not BOT_TOKEN:
        st.error("Please enter BOT TOKEN")
        st.stop()

    if ticket_file is None or master_file is None:
        st.error("Please upload both Excel files")
        st.stop()

    tickets = pd.read_excel(ticket_file)
    engineers = pd.read_excel(master_file)

    sent_log = []

    progress_bar = st.progress(0)
    total = len(engineers)

    for i, eng_row in engineers.iterrows():

        current_time = datetime.now().strftime("%d %B %Y | %I:%M %p")

        taluk = eng_row["Taluk"]
        chat_id = eng_row["Telegram_ID"]
        eng_name = eng_row.get("Engineer Name", "")

        if pd.isna(chat_id):
            continue

        taluk_tickets = tickets[tickets["Taluk"] == taluk]

        if not taluk_tickets.empty:

            msg = (
                f"🕒 <b>Date & Time:</b> {current_time}\n\n"
                "🚨 <b>OPEN TICKET ALERT</b>\n\n"
                f"👨‍🔧 <b>Engineer:</b> {eng_name}\n"
                f"📍 <b>Taluk:</b> {taluk}\n"
                f"📊 <b>Total Tickets:</b> {len(taluk_tickets)}\n\n"
            )

            for _, row in taluk_tickets.iterrows():
                msg += (
                    f"\n🆔 <b>Ticket:</b> {row.get('Ticket Number','')}\n"
                    f"📂 <b>SubCat:</b> {row.get('Sub Category','')}\n"
                    f"⚠ <b>Priority:</b> {row.get('Priority','')}\n"
                    "➖➖➖➖➖➖➖➖➖\n"
                )

            tickets_count = len(taluk_tickets)

        else:
            msg = (
                f"🕒 <b>Date & Time:</b> {current_time}\n\n"
                "✅ <b>No Open Tickets</b>\n\n"
                f"👨‍🔧 <b>Engineer:</b> {eng_name}\n"
                f"📍 <b>Taluk:</b> {taluk}\n"
            )
            tickets_count = 0

        response = send_telegram(chat_id, msg)

        delivery_status = "SUCCESS" if response.status_code == 200 else "FAILED"

        sent_log.append({
            "Date_Time": datetime.now(),
            "Engineer Name": eng_name,
            "Taluk": taluk,
            "Telegram_ID": chat_id,
            "Tickets_Count": tickets_count,
            "HTTP_Status": response.status_code,
            "Delivery_Status": delivery_status
        })

        progress_bar.progress((i+1)/total)

    log_df = pd.DataFrame(sent_log)

    st.success("Alerts Sent Successfully ✅")
    st.dataframe(log_df)

    # Download option
    file_name = f"Sent_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    log_df.to_excel(file_name, index=False)

    with open(file_name, "rb") as f:
        st.download_button("📥 Download Log File", f, file_name)