import streamlit as st
import pandas as pd
import datetime
import random
import io
import urllib.parse
from PIL import Image as PILImage
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="NFP | Attendance Generator",
    page_icon="📊",
    layout="wide"
)

# --- BRAND CONSTANTS ---
BRAND_NAME = "NazeerFinPro"
BRAND_SHORT = "NFP"
BRAND_COLOR = "#003366" # Professional Navy Blue

# --- CSS FOR PROFESSIONAL LOOK ---
st.markdown(f"""
    <style>
    .main {{
        padding-top: 0rem;
    }}
    .stButton>button {{
        width: 100%;
        background-color: {BRAND_COLOR};
        color: white;
        font-weight: bold;
        border-radius: 5px;
    }}
    .stButton>button:hover {{
        background-color: #002244;
        color: white;
    }}
    .brand-header {{
        color: {BRAND_COLOR};
        font-size: 36px;
        font-weight: bold;
        text-align: left;
        margin-bottom: 0px;
        line-height: 1.2;
        padding-top: 10px;
    }}
    .brand-sub {{
        color: #666;
        font-size: 16px;
        text-align: left;
        margin-top: 0px;
        margin-bottom: 20px;
    }}
    .footer {{
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f1f1f1;
        color: #333;
        text-align: center;
        padding: 10px;
        font-size: 12px;
        z-index: 999;
    }}
    div[data-testid="column"] {{
        display: flex;
        align-items: center; 
    }}
    </style>
    """, unsafe_allow_html=True)

# --- CONSTANTS FOR GENERATOR ---
SHIFT_STRING = "(0900:1800)"
STANDARD_HOURS_PER_DAY = 9

# --- HELPER FUNCTIONS (Backend Logic) ---

def create_natural_time(year, month, base_hour, is_arrival):
    """Generates a natural-looking time string."""
    if is_arrival:
        minute = random.randint(-5, 10)
    else:
        minute = random.randint(0, 10)
    
    try:
        base_time = datetime.datetime(year, month, 1, base_hour, 0)
        final_time = base_time + datetime.timedelta(minutes=minute)
        return final_time.strftime("%H:%M")
    except ValueError:
        return "00:00"

def distribute_overtime(required_ot, num_working_days):
    """Distributes required OT hours randomly."""
    if num_working_days == 0:
        return []
        
    ot_hours_list = [0] * num_working_days
    hours_distributed = 0
    
    max_attempts = required_ot * 5 
    attempts = 0
    
    while hours_distributed < required_ot and attempts < max_attempts:
        attempts += 1
        day_index = random.randint(0, num_working_days - 1)
        ot_to_add = random.choice([1, 1, 2])
        
        if hours_distributed + ot_to_add > required_ot:
            ot_to_add = required_ot - hours_distributed
            
        if ot_hours_list[day_index] < 2:
           ot_to_add_today = min(ot_to_add, 2 - ot_hours_list[day_index])
           ot_hours_list[day_index] += ot_to_add_today
           hours_distributed += ot_to_add_today
        
        if all(ot >= 2 for ot in ot_hours_list):
            if hours_distributed < required_ot:
                remaining = required_ot - hours_distributed
                for _ in range(remaining):
                    day_index = random.randint(0, num_working_days - 1)
                    ot_hours_list[day_index] += 1
                hours_distributed = sum(ot_hours_list)
            break
            
    return ot_hours_list

def generate_attendance_file(input_df, target_month, target_year, holidays_dict, company_name_input):
    output = io.BytesIO()
    month_year_str = f"{datetime.date(target_year, target_month, 1).strftime('%B %Y').upper()}"

    title_font = Font(name='Calibri', size=14, bold=True)
    header_font = Font(name='Calibri', size=11, bold=True)
    normal_font = Font(name='Calibri', size=11)
    center_align = Alignment(horizontal='center', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')
    link_font = Font(name='Calibri', size=11, color="0000FF", underline="single")
    thin_side = Side(border_style='thin', color='000000')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    index_data = []

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        index_ws = writer.book.create_sheet(title="Index", index=0)
        
        progress_bar = st.progress(0)
        total_emps = len(input_df)
        
        for i, (_, employee) in enumerate(input_df.iterrows()):
            progress_bar.progress((i + 1) / total_emps)

            emp_code = employee['CODE']
            emp_name = employee['NAME']
            req_ot = int(employee['Overtime Hours'])
            s_no = employee['S#']
            
            try:
                num_absent = int(employee['ABSENT DAYS']) if pd.notna(employee['ABSENT DAYS']) else 0
            except ValueError:
                num_absent = 0
                
            safe_name = str(emp_name).replace(":", "").replace("/", "")
            sheet_name = f"{emp_code}_{safe_name}"[:31]
            ws = writer.book.create_sheet(title=sheet_name)
            
            # Header Data
            header_data = [
                ["Company Name:", company_name_input],
                ["Report Title:", f"ATTENDANCE SHEETS FOR THE MONTH OF {month_year_str}"],
                ["Employee Name:", emp_name],
                ["Employee Code:", emp_code]
            ]
            
            try:
                next_month = datetime.date(target_year, target_month, 28) + datetime.timedelta(days=4)
                last_day_of_month = next_month - datetime.timedelta(days=next_month.day)
                num_days_in_month = last_day_of_month.day
            except ValueError:
                num_days_in_month = 30
                
            working_days_in_month = []
            full_month_data = []
            sundays = 0
            holidays_found = 0
            
            for day_num in range(1, num_days_in_month + 1):
                current_date = datetime.date(target_year, target_month, day_num)
                is_sunday = current_date.weekday() == 6
                is_holiday = current_date in holidays_dict
                
                if is_sunday:
                    sundays += 1
                elif is_holiday:
                    holidays_found += 1
                else:
                    working_days_in_month.append(current_date)
            
            num_working_days = len(working_days_in_month)
            absent_days = set()
            if num_absent > 0 and num_absent <= num_working_days:
                absent_days = set(random.sample(working_days_in_month, num_absent))
                
            index_data.append({
                "S. No": s_no,
                "CODE": emp_code,
                "Name": emp_name,
                "SheetName": sheet_name,
                "Absent": num_absent,
                "OT Hours": req_ot
            })
            
            working_days_with_attendance = [day for day in working_days_in_month if day not in absent_days]
            ot_schedule = distribute_overtime(req_ot, len(working_days_with_attendance))
            
            work_day_counter = 0
            
            for day_num in range(1, num_days_in_month + 1):
                current_date = datetime.date(target_year, target_month, day_num)
                date_str = current_date.strftime("%d-%b-%y")
                is_sunday = current_date.weekday() == 6
                holiday_name = holidays_dict.get(current_date)
                
                row = [date_str, SHIFT_STRING, "", "", "", ""]
                
                if is_sunday:
                    row[5] = "SUNDAY"
                elif holiday_name:
                    row[5] = holiday_name
                elif current_date in working_days_with_attendance:
                    ot_hours = ot_schedule[work_day_counter]
                    row[2] = create_natural_time(target_year, target_month, 9, True)
                    row[3] = create_natural_time(target_year, target_month, 18 + ot_hours, False)
                    row[4] = ot_hours if ot_hours > 0 else ""
                    row[5] = "On Time"
                    work_day_counter += 1
                elif current_date in working_days_in_month:
                    row[5] = "Absent"
                
                full_month_data.append(row)
                
            # Footing Logic
            total_present_days = work_day_counter 
            total_std_hours = (work_day_counter) * STANDARD_HOURS_PER_DAY
            total_ot_hours = sum(ot_schedule)
            total_payable_hours = total_std_hours + total_ot_hours
            
            footing_data = [
                ["SUMMARY:", ""],
                ["Total Days in Month", num_days_in_month],
                ["Sundays", sundays],
                ["Gazetted Holidays", holidays_found],
                ["Total Present Days", total_present_days],
                ["Absent", num_absent],
                ["Over Time Hrs.", total_ot_hours],
                [],
                ["Total Standard Hours", total_std_hours, f"({work_day_counter} Days x {STANDARD_HOURS_PER_DAY} Hrs)"],
                ["Total OT Hours", total_ot_hours, f"(Sum of OT HRS)"],
                ["Total Payable Hours", total_payable_hours]
            ]
            
            # --- WRITING & FORMATTING ---
            
            # 1. Header Formatting
            ws.cell(row=1, column=1, value="Company Name:").font = title_font
            ws.cell(row=1, column=2, value=company_name_input).font = title_font 
            ws.merge_cells('B1:F1') 

            for r_idx, row_val in enumerate(header_data[1:], 2):
                ws.cell(row=r_idx, column=1, value=row_val[0]).font = header_font
                ws.cell(row=r_idx, column=2, value=row_val[1]).font = normal_font
                ws.merge_cells(f'B{r_idx}:F{r_idx}') 

            # 2. Data Table
            data_start_row = 6
            table_headers = ["DATE", "SHIFT G", "TIME IN", "TIME OUT", "OT HRS", "REMARKS"]
            
            for c_idx, val in enumerate(table_headers, 1):
                cell = ws.cell(row=data_start_row, column=c_idx, value=val)
                cell.font = header_font
                cell.border = thin_border
                cell.alignment = center_align
                
            for r_idx, row_val in enumerate(full_month_data, data_start_row + 1):
                for c_idx, val in enumerate(row_val, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = normal_font
                    cell.border = thin_border
                    cell.alignment = center_align
                    if c_idx == 5: cell.alignment = right_align

            # 3. Footing
            footing_start_row = data_start_row + len(full_month_data) + 2
            ws.cell(row=footing_start_row, column=1, value="SUMMARY:").font = header_font
            ws.cell(row=footing_start_row, column=1).border = thin_border
            
            for r_idx, row_val in enumerate(footing_data[1:], footing_start_row + 1):
                for c_idx, val in enumerate(row_val, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = normal_font
                    cell.border = thin_border
                    if c_idx == 1: cell.font = header_font
                    if c_idx > 1: cell.alignment = right_align

            # 4. Dimensions & Print
            widths = [15, 15, 12, 12, 10, 20]
            for i, w in enumerate(widths):
                ws.column_dimensions[get_column_letter(i+1)].width = w
                
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1

        # --- Index Sheet Logic ---
        index_headers = ["S. No", "CODE", "Name", "Absent", "OT Hours"]
        for c_idx, val in enumerate(index_headers, 1):
            cell = index_ws.cell(row=1, column=c_idx, value=val)
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = center_align
            
        for r_idx, data in enumerate(index_data, 2):
            c1 = index_ws.cell(row=r_idx, column=1, value=data['S. No'])
            c1.font = normal_font; c1.border = thin_border; c1.alignment = center_align
            c2 = index_ws.cell(row=r_idx, column=2, value=data['CODE'])
            c2.font = normal_font; c2.border = thin_border; c2.alignment = center_align
            c3 = index_ws.cell(row=r_idx, column=3)
            c3.value = f'=HYPERLINK("#\'{data["SheetName"]}\'!A1", "{data["Name"]}")'
            c3.font = link_font; c3.border = thin_border
            c4 = index_ws.cell(row=r_idx, column=4, value=data['Absent'])
            c4.font = normal_font; c4.border = thin_border; c4.alignment = center_align
            c5 = index_ws.cell(row=r_idx, column=5, value=data['OT Hours'])
            c5.font = normal_font; c5.border = thin_border; c5.alignment = center_align
            
        if 'Sheet' in writer.book.sheetnames:
            writer.book.remove(writer.book['Sheet'])

    return output

# --- SIDEBAR: BRANDING & NAVIGATION ---
with st.sidebar:
    # 1. LOGO & HERO IMAGE (Kept the office image here as requested)
    try:
        logo = PILImage.open("nfp_office.jpg") 
        st.image(logo, use_container_width=True)
    except:
        try:
            logo = PILImage.open("Nazeer Fin Pro - NFP.jpg")
            st.image(logo, use_container_width=True)
        except:
            st.header(f"✨ {BRAND_SHORT}")
        
    st.markdown(f"## **{BRAND_NAME}**")
    st.caption("Professional Finance Consultancy")
    st.write("---")
    
    # Modified Layout: Image visible outside/before the expander on the right
    col_about_label, col_about_img = st.columns([3, 1])
    
    with col_about_label:
        st.write("**About Nazeer Ahmed Khan**")
        
    with col_about_img:
        try:
            about_img = PILImage.open("about_nazeer.jpg")
            st.image(about_img, use_container_width=True)
        except:
            pass
            
    with st.expander("View Details"):
        st.write("""
        **Nazeer Ahmed Khan** is the founder of NazeerFinPro (NFP). 
        
        A seasoned Finance Freelance Consultant specializing in:
        - Financial Analysis & Reporting
        - Automation & Data Cleaning (Python/Excel)
        - Business Intelligence Dashboards
        - Corporate Training
        
        *\"Turning chaotic data into clear, actionable financial insights.\"*
        """)
    
    st.write("---")

    # APP SETTINGS
    st.header("⚙️ Generator Settings")
    
    # --- COMPANY NAME INPUT ---
    st.info("👇 Enter Company Details")
    company_name = st.text_input("Company Name", value="ABC COMPANY")
    
    target_date = st.date_input("Select Month & Year", datetime.date(2025, 8, 1))
    selected_month = target_date.month
    selected_year = target_date.year
    
    st.subheader("🎉 Gazetted Holidays")
    
    if 'holidays' not in st.session_state:
        st.session_state.holidays = [
            {"date": datetime.date(selected_year, 8, 14), "name": "Independence Day"}
        ]
    
    with st.form("add_holiday"):
        default_date_val = datetime.date(selected_year, selected_month, 1)
        h_date = st.date_input("Holiday Date", value=default_date_val)
        h_name = st.text_input("Holiday Name", placeholder="e.g. Eid, Labor Day")
        
        submitted = st.form_submit_button("Add Holiday")
        if submitted:
            if h_name:
                st.session_state.holidays.append({"date": h_date, "name": h_name})
                st.success(f"Added: {h_name} on {h_date}")
            else:
                st.error("Please enter a name for the holiday.")
            
    holidays_dict = {}
    current_month_holidays = []

    if st.session_state.holidays:
        for h in st.session_state.holidays:
            if h['date'].year == selected_year and h['date'].month == selected_month:
                current_month_holidays.append(h)
                if h['date'] in holidays_dict:
                    holidays_dict[h['date']] += f" / {h['name']}"
                else:
                    holidays_dict[h['date']] = h['name']
    
    if current_month_holidays:
        st.write("**Active Holidays for Report:**")
        display_data = [{"Date": h['date'].strftime('%d-%b-%Y'), "Name": h['name']} for h in current_month_holidays]
        st.dataframe(display_data, hide_index=True, use_container_width=True)
    else:
        st.info("No holidays added for this specific month.")

    if st.button("Clear All Holidays"):
        st.session_state.holidays = []
        st.rerun()

# --- MAIN CONTENT ---

# --- HEADER SECTION ---
col_title, col_logo = st.columns([5, 1])

with col_title:
    st.markdown(f"<div class='brand-header'>{BRAND_NAME} Tool Suite</div>", unsafe_allow_html=True)
    st.markdown("<div class='brand-sub'>Advanced Financial Solutions & Automation</div>", unsafe_allow_html=True)

with col_logo:
    try:
        # --- RESTORED ORIGINAL LOGO HERE ---
        header_logo = PILImage.open("logo.jpg")
        st.image(header_logo, width=150) 
    except:
        pass 


# Tabs for sections
tab1, tab2, tab3 = st.tabs(["🏢 Attendance Generator", "📝 Consultancy Blog", "📞 Contact NFP"])

# --- TAB 1: THE GENERATOR APP ---
with tab1:
    st.subheader("Auto-Generate Attendance Sheets")
    st.info("Upload your employee data file (`data.xlsx`) to generate payroll-ready Excel sheets with natural time variations.")
    
    uploaded_file = st.file_uploader("Upload Input File", type=['xlsx'])

    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.success("File loaded!")
            with st.expander("View Input Data"):
                st.dataframe(df.head())
                
            if st.button("🚀 Generate & Download Report", type="primary"):
                with st.spinner("Processing data..."):
                    excel_data = generate_attendance_file(df, selected_month, selected_year, holidays_dict, company_name)
                    st.success("Done! Your file is ready.")
                    
                    file_name = f"NFP_Attendance_{target_date.strftime('%B_%Y')}.xlsx"
                    st.download_button(
                        label="📥 Download Excel File",
                        data=excel_data.getvalue(),
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"Error: {e}")

# --- TAB 2: BLOG PLACEHOLDER ---
with tab2:
    st.subheader("📰 NFP Financial Insights")
    st.write("Welcome to the NazeerFinPro blog. Here we share insights on financial management and automation.")
    
    col1, col2 = st.columns(2)
    with col1:
        try:
            st.image("https://images.unsplash.com/photo-1554224155-8d04cb21cd6c?w=800", caption="Financial Analysis")
        except:
            pass
        st.markdown("#### 5 Ways to Automate Your Payroll")
        st.write("Stop doing manual data entry. Learn how Python can save you 20 hours a week...")
        st.button("Read More", key="b1")
        
    with col2:
        try:
            st.image("https://images.unsplash.com/photo-1460925895917-afdab827c52f?w=800", caption="Data Visualization")
        except:
             pass
        st.markdown("#### The Power of Visual Data")
        st.write("Why your stakeholders ignore your spreadsheets and how to fix it with dashboards...")
        st.button("Read More", key="b2")

# --- TAB 3: CONTACT ---
with tab3:
    st.subheader("🤝 Work with NazeerFinPro")
    st.write("Ready to automate your financial processes? Let's connect.")
    
    c1, c2 = st.columns([1, 2])
    with c1:
        try:
            # --- UPDATED IMAGE HERE ---
            profile = PILImage.open("Nazeer Fin Pro - NFP.jpg") 
            st.image(profile, width=200)
        except:
            st.info("[Profile Photo Placeholder]")
            
    with c2:
        st.markdown("""
        **Nazeer Ahmed Khan** *Finance Consultant & Automation Expert*
        
        📧 **Email:** [nazeerfinpro@gmail.com](mailto:nazeerfinpro@gmail.com)  
        📱 **Phone / WhatsApp:** [+92 333 3126614](https://wa.me/923333126614)  
        🌐 **Website:** [LinkedIn Company Page](https://www.linkedin.com/company/nazeer-fin-pro)  
        """)
        
        st.write("---")
        st.write("**Send a Message:**")
        contact_msg = st.text_area("Message content:", placeholder="Hi Nazeer, I need help with...", label_visibility="collapsed")
        
        if st.button("Prepare Message"):
            if contact_msg:
                subject = "Inquiry from NFP Tool Suite"
                encoded_body = urllib.parse.quote(contact_msg)
                encoded_subject = urllib.parse.quote(subject)
                
                mailto_url = f"mailto:nazeerfinpro@gmail.com?subject={encoded_subject}&body={encoded_body}"
                whatsapp_url = f"https://wa.me/923333126614?text={encoded_body}"
                
                st.success("Message Prepared! Choose how to send it:")
                
                st.markdown(f"""
                <div style="display: flex; gap: 10px; margin-top: 10px;">
                    <a href="{mailto_url}" target="_blank" style="
                        padding: 10px 20px;
                        background-color: #003366;
                        color: white;
                        text-decoration: none;
                        border-radius: 5px;
                        font-weight: bold;
                        border: 1px solid #002244;">
                        📧 Send via Email
                    </a>
                    <a href="{whatsapp_url}" target="_blank" style="
                        padding: 10px 20px;
                        background-color: #25D366;
                        color: white;
                        text-decoration: none;
                        border-radius: 5px;
                        font-weight: bold;
                        border: 1px solid #128C7E;">
                        💬 Send via WhatsApp
                    </a>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.warning("⚠️ Please write a message first.")

# Footer
st.markdown(f"<div class='footer'>© 2025 {BRAND_NAME}. All Rights Reserved. Powered by Python & Streamlit.</div>", unsafe_allow_html=True)