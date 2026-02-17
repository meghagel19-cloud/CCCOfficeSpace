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
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.exceptions import APIError

try:
    # Load credentials from Streamlit Secrets
    creds_dict = st.secrets["gcp_service_account"]

    # Fix private_key newlines (required for Streamlit Secrets)
    if "\\n" in creds_dict["private_key"]:
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

    # Create Google credentials object
    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        creds_dict,
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ],
    )

    # Authorize gspread
    gc = gspread.authorize(creds)

    # Open your spreadsheet
    SPREADSHEET_ID = creds_dict["spreadsheet_id"]
    sh = gc.open_by_key(SPREADSHEET_ID)
    worksheet = sh.sheet1

    st.success("✅ Connected to Google Sheets successfully!")

except KeyError as ke:
    st.error(f"Missing key in Secrets.toml: {ke}")
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

# --- Mapping of desk_labels to their default team member ---
# DEFAULT_DESK_ASSIGNMENTS = {
#     "Branch Head Office": "PS",
#     "Tech Head Office": "PL",
#     "Cyber Desk 1": "MF",
#     "Cyber Desk 2": "PW",
#     "Admin Desk": "JR",
#     "Temp Desk": "MH",
#     "DCIS Desk 1": "AK",
#     "DCIS Desk 2": "MC",
#     "Core Desk 1": "DG",
#     "Core Desk 2": "GM",
#     "Open Space Desk": "CG"
# }

# === Load all bookings from sheet every run ===
bookings = {}
try:
    for rec in worksheet.get_all_records():
        # Clean up date, strip leading apostrophe if present
        raw_date = rec.get("Date")
        if isinstance(raw_date, str):
            date_str = raw_date.lstrip("'")
        else:
            date_str = str(raw_date)
        desk_name = rec.get("Desk")
        user = rec.get("Booked By")
        if date_str and desk_name and user and desk_name in desk_labels:
            idx = desk_labels.index(desk_name) + 1
            key = f"{date_str}_desk{idx}"
            bookings[key] = user
except Exception as e:
    st.error(f"Could not load existing bookings from Google Sheets: {e}")

# === Callback to write a single booking to Google Sheets ===
# === Callback to write a single booking to Google Sheets ===
def write_booking(key):
    """
    Write a booking to Google Sheets when a dropdown changes.
    Only appends if the value has changed.
    """
    val = st.session_state[key]
    prev = bookings.get(key)
    if val == prev:
        return  # no change, do nothing

    # Extract date and desk index from key
    try:
        date_str, desk = key.split("_")
        idx = int(desk.replace("desk", ""))
        # Append the booking row
        worksheet.append_row([date_str, desk_labels[idx-1], val])
        bookings[key] = val
        st.success(f"Booked {val} for {desk_labels[idx-1]} on {date_str}")
    except ValueError:
        st.error(f"Invalid key format: {key}")
    except Exception as e:
        st.error(f"Failed to save booking: {e}")

# === Calendar Rendering & Dropdowns ===

today = datetime.today()
# Auto-scroll for the configured year
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

# Render full year (Jan–Dec)
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

                            # Initialize session state only if key is missing
                            if key not in st.session_state:
                                st.session_state[key] = bookings.get(key, "")

                            # Set default selection safely
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
