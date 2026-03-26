import streamlit as st
import pandas as pd
import subprocess
import os

# 1. ตั้งค่าหน้าเว็บ
st.set_page_config(page_title="Log-Sleuth: Smart Analyzer", layout="wide")

# --- ส่วนของฟังก์ชันวิเคราะห์ (Logic) ---
def translate_error(row):
    msg = str(row['Message']).lower()
    eid = str(row['Id'])
    
    # 1. ปัญหา Excel (ID 1000)
    if '1000' in eid and 'excel.exe' in msg:
        return "❌ [แอปฯ ค้าง/เด้ง] Excel ปิดตัวกะทันหัน มักเกิดจากไฟล์เสียหรือ Add-in มีปัญหา"
    
    # 2. ปัญหา Network / Domain (ID 5719 / 1129)
    elif eid in ['5719', '1129']:
        return "🌐 [การเชื่อมต่อบริษัท] เครื่องติดต่อ Server บริษัท (PTG) ไม่ได้ ตรวจสอบ VPN หรือสายแลน"
    
    # 3. ปัญหา HP Service (ID 7031)
    elif '7031' in eid and 'hp insights' in msg:
        return "🛠️ [บริการ HP ขัดข้อง] โปรแกรมช่วยซ่อมของ HP หยุดทำงานชั่วคราว ไม่มีผลกับการใช้งานทั่วไป"
    
    # 4. ปัญหา .NET Runtime (ID 1022)
    elif '1022' in eid and '.net runtime' in msg:
        return "💻 [ระบบพื้นฐาน Windows] มีปัญหาการโหลดโปรแกรมบางตัว แนะนำให้ลอง Restart เครื่อง"
    
    # 5. ปัญหา Wi-Fi Adapter (ID 10317)
    elif '10317' in eid:
        return "📶 [ฮาร์ดแวร์ Wi-Fi] การ์ด Wi-Fi มีปัญหาตอนสลับโหมดประหยัดไฟ แนะนำให้เสียบปลั๊กชาร์จ"

    # 6. กรณีอื่นๆ ที่ไม่ระบุไว้
    return "🔍 [ตรวจสอบเพิ่มเติม] พบข้อผิดพลาดทั่วไป โปรดอ่านรายละเอียดใน Message"

# --- ส่วนของ Sidebar (แถบเครื่องมือด้านข้าง - เวอร์ชันปลอดภัย 100%) ---
with st.sidebar:
    st.header("🔍 Diagnostic Tools")
    st.write("ตรวจสอบสถานะ (ไม่แก้ไขระบบ)")

    # ปุ่มที่ 1: เช็คการเชื่อมต่อ Domain
    if st.button('📡 Check Domain Sync'):
        with st.spinner('กำลังตรวจสอบการเชื่อมต่อกับ Domain...'):
            # ลอง Test Connection กับ Domain (ในที่นี้สมมติชื่อ PTG ตาม Log ของคุณ)
            # เราใช้คำสั่ง Test-Connection ซึ่งปลอดภัยกว่าการสั่ง Restart
            nav_test = subprocess.run(["powershell.exe", "-Command", "Test-Connection -ComputerName PTG -Count 1 -Quiet"], capture_output=True, text=True)
            
            if "True" in nav_test.stdout:
                st.success("✅ เชื่อมต่อกับ Domain PTG ได้ปกติ!")
            else:
                st.error("❌ ไม่สามารถติดต่อ Domain ได้ (ตรวจสอบ VPN/สายแลน)")

    # ปุ่มที่ 2: เช็คสถานะ Network เบื้องต้น
    if st.button('🌍 Test Internet Latency'):
        with st.spinner('กำลังทดสอบความเร็วตอบสนอง...'):
            ping_test = subprocess.run(["powershell.exe", "-Command", "ping 8.8.8.8 -n 1"], capture_output=True, text=True)
            if "Reply from" in ping_test.stdout:
                # ดึงตัวเลข ms ออกมา
                import re
                ms_match = re.search(r'time[=<](\d+)ms', ping_test.stdout)
                ms_val = int(ms_match.group(1)) if ms_match else 0
                
                if ms_val < 50:
                    quality = "🚀 ยอดเยี่ยม (Excellent)"
                elif ms_val < 150:
                    quality = "👍 ปกติ (Good)"
                else:
                    quality = "🐢 ช้า/หน่วง (Laggy)"
                
                st.success(f"🌐 อินเทอร์เน็ตใช้งานได้: {quality}")
                st.metric("Latency", f"{ms_val} ms")
            else:
                st.error("⚠️ ติดต่อภายนอกไม่ได้ (Internet Down)")

    # ปุ่มที่ 3: เช็คสุขภาพ Hardware เบื้องต้น (Power & Battery)
    if st.button('🔋 Check Power Health'):
        with st.spinner('กำลังดึงข้อมูลพลังงาน...'):
            # ดึงสถานะแบตเตอรี่ พร้อมสถานะการเสียบปลั๊ก (BatteryStatus + PowerOnline)
            pwr_cmd = "Get-CimInstance -ClassName Win32_Battery | Select-Object EstimatedChargeRemaining, BatteryStatus"
            # เช็คว่าเสียบปลั๊กอยู่ไหมจากอีกที่หนึ่ง (Win32_PortableBattery หรือ BatteryStatus)
            pwr_test = subprocess.run(["powershell.exe", "-Command", pwr_cmd], capture_output=True, text=True)
            
            st.write("📊 **สรุปสถานะพลังงานปัจจุบัน:**")
            
            raw_output = pwr_test.stdout.strip()
            if raw_output:
                lines = raw_output.split('\n')
                if len(lines) >= 3:
                    values = lines[2].split()
                    charge = int(values[0])
                    status_id = values[1]
                    
                    # แก้บั๊ก: ถ้าแบตเกิน 95% และระบบรายงานว่า Discharging (2) 
                    # ให้สันนิษฐานว่าเป็นระบบถนอมแบตเตอรี่ (Fully Charged/Idle)
                    if charge >= 98 and status_id == "2":
                        human_status = "🔌 เสียบปลั๊กอยู่ - แบตเตอรี่เต็ม (Fully Charged)"
                        st.metric(label="ปริมาณแบตเตอรี่คงเหลือ", value="100%", delta="Fully Charged")
                        st.success(f"**สถานะปัจจุบัน:** {human_status}")
                    else:
                        status_map = {
                            "1": "❓ ไม่ระบุสถานะ",
                            "2": "🔋 กำลังใช้แบตเตอรี่ (Discharging)",
                            "3": "🔌 เสียบปลั๊กอยู่ - แบตเต็ม (Fully Charged)",
                            "6": "⚡ กำลังชาร์จไฟ (Charging)",
                            "10": "🔌 เสียบปลั๊กอยู่ (AC Power)"
                        }
                        human_status = status_map.get(status_id, f"รหัสสถานะ: {status_id}")
                        st.metric(label="ปริมาณแบตเตอรี่คงเหลือ", value=f"{charge}%")
                        
                        if status_id in ["2", "4", "5"]:
                            st.warning(f"**สถานะปัจจุบัน:** {human_status}")
                            st.info("💡 **คำแนะนำ:** หาก Wi-Fi หลุดบ่อย ลองเสียบชาร์จเพื่อเพิ่มประสิทธิภาพครับ")
                        else:
                            st.success(f"**สถานะปัจจุบัน:** {human_status}")
            else:
                st.success("💻 **สถานะ:** เครื่อง Desktop PC (ใช้ไฟบ้านปกติ)")

    if st.button('🌐 Check IP Address'):
        with st.spinner('กำลังตรวจสอบที่อยู่ IP...'):
            # แก้ไข: กรองเอาเฉพาะ Adapter ที่กำลังต่อเน็ตอยู่จริงๆ (Status = Up)
            local_ip_cmd = "Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway -ne $null } | Select-Object -ExpandProperty IPv4Address | Select-Object -ExpandProperty IPAddress"
            local_ip_test = subprocess.run(["powershell.exe", "-Command", local_ip_cmd], capture_output=True, text=True)
            
            # ล้างค่าว่างและตัวหนังสือส่วนเกิน
            local_ip = local_ip_test.stdout.strip().split('\n')[0] if local_ip_test.stdout.strip() else "หาไม่พบ (เช็คสายแลน/Wi-Fi)"

            # เช็ค Public IP
            import urllib.request
            try:
                public_ip = urllib.request.urlopen('https://ident.me', timeout=3).read().decode('utf8')
            except:
                public_ip = "Offline (ติดต่อภายนอกไม่ได้)"

            st.write("📍 **สรุปที่อยู่ Network:**")
            col_a, col_b = st.columns(2)
            with col_a:
                st.metric("Local IP (ในเครื่อง)", local_ip)
            with col_b:
                st.metric("Public IP (ขาออก)", public_ip)

            # วิเคราะห์วงเน็ตแบบเจาะจง
            if local_ip.startswith("10.") or local_ip.startswith("172.16."):
                st.success("🏢 **วิเคราะห์:** เชื่อมต่อผ่านวง Network องค์กร (Corporate)")
            elif local_ip.startswith("192.168."):
                st.warning("🏠 **วิเคราะห์:** เชื่อมต่อผ่าน Router บ้าน/Hotspot")

    st.divider()
    st.write("⚠️ *หมายเหตุ: ทุกปุ่มเป็นการอ่านค่าเท่านั้น ไม่มีการเปลี่ยนแปลงค่าใดๆ ในเครื่อง*")

# --- ส่วนหน้าจอหลัก (Main UI) ---
st.title("🔍 Log-Sleuth Dashboard")
st.subheader("ระบบวิเคราะห์ Error ย้อนหลังสำหรับ Service Desk")

if st.button('🚀 เริ่มการสแกนระบบ (Run Diagnostic)'):
    with st.spinner('กำลังดึงข้อมูลจาก Event Viewer...'):
        # สั่งรัน PowerShell Script
        process = subprocess.run(["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", "get_logs.ps1"], capture_output=True)
        
        if os.path.exists("current_logs.csv"):
            data = pd.read_csv("current_logs.csv")
            
            if not data.empty:
                # รัน Logic การวิเคราะห์
                data['Quick Solution'] = data.apply(translate_error, axis=1)
                
                st.success(f"สแกนเสร็จสิ้น! พบปัญหาที่ควรตรวจสอบ {len(data)} รายการ")
                
                # แสดงกราฟและสรุป
                col1, col2 = st.columns(2)
                with col1:
                    st.write("📊 สถิติแยกตามประเภท Log")
                    st.bar_chart(data['LogName'].value_counts())
                
                with col2:
                    st.write("💡 วิเคราะห์ภาพรวมอัตโนมัติ")
                    top_error = data['Id'].mode()[0]
                    st.warning(f"พบ Error ID: {top_error} บ่อยที่สุดในเครื่องนี้")

                # ส่วนที่แสดงตารางข้อมูล
                st.write("📋 **สรุปรายการปัญหาที่พบ (เรียงตามเวลา):**")
                # เปลี่ยนชื่อคอลัมน์ให้เป็นภาษาไทยเพื่อความสวยงาม
                display_df = data[['TimeCreated', 'LogName', 'Id', 'Quick Solution', 'Message']].copy()
                display_df.columns = ['วัน-เวลาที่เกิดเหตุ', 'ประเภท Log', 'รหัส Error ID', 'สรุปอาการ/วิธีแก้', 'รายละเอียดจากระบบ']
                
                # แสดงผลตารางแบบเต็มความกว้างหน้าจอ
                st.dataframe(display_df, use_container_width=True)                
            else:
                st.balloons()
                st.success("ไม่พบ Error ใดๆ ใน 24 ชั่วโมงที่ผ่านมา ระบบปกติดีครับ!")
        else:
            st.error("ไม่สามารถสร้างรายงานได้ โปรดลองรันด้วยสิทธิ์ Administrator")

st.markdown("---")
st.caption("Developed by Brooklyn | Service Desk Automation Project v2.0")