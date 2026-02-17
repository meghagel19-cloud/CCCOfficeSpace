import streamlit as st
import calendar
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.exceptions import APIError

# === Config ===
YEAR = 2026

# === App Setup ===
st.set_page_config(page_title=f"Desk Booking – {YEAR}", layout="wide")
st.title(f"📅 Office Desk Booking – {YEAR}")

# === Google Sheets Setup ===
worksheet = None  # placeholder to avoid scope errors
try:
    creds_dict = st.secrets["gcp_service_account"]

    # Fix private_key newlines
    if "\\n" in creds_dict["private_key"]:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        creds_dict,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )
    gc = gspread.authorize(creds)
    SPREADSHEET_ID = creds_dict["spreadsheet_id"]
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.sheet1
    st.success("✅ Connected to Google Sheets successfully!")

except KeyError:
    st.error("Secrets key 'gcp_service_account' not found. Check your Streamlit Secrets.")
except APIError as api_err:
    st.error(f"Google Sheets API error: {api_err}")
except Exception as e:
    st.error(f"Failed to connect to Google Sheets: {e}")

# === Desk & Team Setup ===
desk_labels = [
    "Branch Head Office",
    "Tech Head Office",
    "Cyber Desk 1",
    "Cyber Desk 2",
    "Core Desk 1",
    "Core Desk 2",
    "DCIS Desk 1",
    "DCIS Desk 2",
    "Admin Desk",
    "Temp Desk",
    "Open Space Desk",
]
team_members = [" ","FREE", "PS", "PL", "MF","PW", "JR", "MH", "AK", "MC", "GM", "DG", "CG"]

# === Load existing bookings ===
bookings = {}
if worksheet:
    try:
        for rec in worksheet.get_all_records():
            raw_date = rec.get("Date")
            date_str = raw_date.lstrip("'") if isinstance(raw_date, str) else str(raw_date)
            desk_name = rec.get("Desk")
            user = rec.get("Booked By")
            if date_str and desk_name and user and desk_name in desk_labels:
                idx = desk_labels.index(desk_name) + 1
                key = f"{date_str}_desk{idx}"
                bookings[key] = user
    except Exception as e:
        st.error(f"Could not load existing bookings from Google Sheets: {e}")

# === Booking callback ===
def write_booking(key):
    global worksheet
    if not worksheet:
        st.error("Google Sheets not connected. Cannot save booking.")
        return

    val = st.session_state[key]
    prev = bookings.get(key)
    if val == prev:
        return  # no change

    try:
        date_str, desk = key.split("_")
        idx = int(desk.replace("desk", ""))
        worksheet.append_row([date_str, desk_labels[idx-1], val])
        bookings[key] = val
        st.success(f"Booked {val} for {desk_labels[idx-1]} on {date_str}")
    except Exception as e:
        st.error(f"Failed to save booking: {e}")

# === Calendar Rendering & Dropdowns ===
today = datetime.today()
if today.year == YEAR:
    today_str = today.strftime("%Y-%m-%d")
    st.markdown(
        f"""
        <script>
            window.onload = function() {{
                var el = document.getElementsByName("{today_str}")[0];
                if (el) el.scrollIntoView({{behavior:'smooth'}});
            }};
        </script>
        """,
        unsafe_allow_html=True,
    )

for month in range(1, 13):
    cal = calendar.monthcalendar(YEAR, month)
    expanded = (month == today.month and today.year == YEAR)
    with st.expander(f"{calendar.month_name[month]} {YEAR}", expanded=expanded):
        for week in cal:
            cols = st.columns(5)
            for i, day in enumerate(week[:5]):
                with cols[i]:
                    if day:
                        date_str = f"{YEAR}-{month:02d}-{day:02d}"
                        st.markdown(f'<a name="{date_str}"></a>', unsafe_allow_html=True)
                        st.markdown(f"### {calendar.day_abbr[i]} {day}")

                        for idx, desk_name in enumerate(desk_labels, start=1):
                            key = f"{date_str}_desk{idx}"
                            # Initialize session state only if missing
                            if key not in st.session_state:
                                st.session_state[key] = bookings.get(key, "")

                            default = st.session_state.get(key, "")
                            idx_default = team_members.index(default) if default in team_members else 0

                            st.selectbox(
                                label=desk_name,
                                options=team_members,
                                index=idx_default,
                                key=key,
                                on_change=write_booking,
                                args=(key,),
                                label_visibility="visible"
                            )
                    else:
                        st.markdown(" ")

