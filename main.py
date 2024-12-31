import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import datetime
import openpyxl
from PyQt5 import QtWidgets, QtCore
import sys

DELAY_UPDATE_INPUT = 0.5
DELAY_EACH_STEP = 15

USERNAME = "ad_huyta"
PASSWORD = "anhhuy123@"

def get_current_hour():
    now = datetime.datetime.now()
    return now.hour

def get_current_minute():
    now = datetime.datetime.now()
    return now.minute

def get_current_date(format="%Y%m%d"):
    return datetime.datetime.now().strftime(format)

# Hàm để đăng nhập và thực hiện nhập thông tin
def auto_login_and_input(current_hour, current_minute, xlsx_folder):
    try:
        processing_hour = current_hour + 1

        if processing_hour == 0:
            processing_hour = 24
        
        hour_in_xlsx = f"{processing_hour}_{current_minute}"

        current_date = get_current_date()
        xlsx_file = os.path.join(xlsx_folder, f"{current_date}.xlsx")

        workbook = openpyxl.load_workbook(xlsx_file)
        sheet = workbook["Thông số"]

        row_data = None
        for row in sheet.iter_rows(min_row=2):
            if row[0].value == hour_in_xlsx:
                row_data = row
                break

        if row_data is None:
            return False
        
        print(f"[{get_current_date(format='%Y-%m-%d')}] Chạy tác vụ {current_hour}_{current_minute}.")

        try:
            # Khởi tạo WebDriver (sử dụng Chrome ở đây)
            driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

            # Mở trang web
            driver.get("https://qltb.capnuochaiphong.com.vn/nhat-ky-01")

            # Wait for the username field to be present
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "tbTaiKhoan"))
            )

            # Tìm và nhập thông tin đăng nhập
            username_field = driver.find_element(By.ID, "tbTaiKhoan")  # Cập nhật NAME phù hợp
            password_field = driver.find_element(By.ID, "tbMatKhau")  # Cập nhật NAME phù hợp
            
            username_field.send_keys(USERNAME)
            password_field.send_keys(PASSWORD)
            password_field.send_keys(Keys.RETURN)  # Nhấn Enter
            print("Đăng nhập thành công.")
            
            # Đợi trang tải sau khi đăng nhập
            time.sleep(DELAY_EACH_STEP)  # Tùy chỉnh thời gian nếu cần
            
            # Tìm và nhập thông tin sau khi đăng nhập
            input_field = driver.find_element(By.XPATH, f"//tr/td[count(//th[.='Giờ'])+1][.='{processing_hour if processing_hour >= 10 else '0' + str(processing_hour)}']/preceding-sibling::td/button")  # Cập nhật NAME phù hợp
            input_field.click()
            print("Đã chọn giờ cần nhập.")

            time.sleep(2)  # Tùy chỉnh thời gian nếu cần

            dien_ap_mba1_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Điện áp MBA1')]]/input")
            dien_ap_mba1_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            dien_ap_mba1_field.send_keys(int(row_data[1].value) if row_data[1].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập điện áp MBA1.")

            dien_ap_mba2_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Điện áp MBA2')]]/input")
            dien_ap_mba2_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            dien_ap_mba2_field.send_keys(int(row_data[2].value) if row_data[2].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập điện áp MBA2.")

            luu_luong_tuyen_1_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Lưu lượng tuyến 1')]]/input")
            luu_luong_tuyen_1_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            luu_luong_tuyen_1_field.send_keys(int(row_data[3].value) if row_data[3].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập lưu lượng tuyến 1.")

            luu_luong_tuyen_2_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Lưu lượng tuyến 2')]]/input")
            luu_luong_tuyen_2_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            luu_luong_tuyen_2_field.send_keys(int(row_data[4].value) if row_data[4].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập lưu lượng tuyến 2.")

            ap_luc_tuyen_1_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Áp lực tuyến 1')]]/input")
            ap_luc_tuyen_1_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            ap_luc_tuyen_1_field.send_keys(float(row_data[5].value) if row_data[5].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập lưu lượng tuyến 2.")

            ap_luc_tuyen_2_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Áp lực tuyến 2')]]/input")
            ap_luc_tuyen_2_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            ap_luc_tuyen_2_field.send_keys(float(row_data[6].value) if row_data[6].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập áp lực tuyến 2.")

            do_dan_dien_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Độ dẫn điện')]]/input")
            do_dan_dien_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            do_dan_dien_field.send_keys(row_data[7].value if row_data[7].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập độ dẫn điện.")

            clo_khu_trung_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Clo Khử trùng')]]/input")
            clo_khu_trung_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            clo_khu_trung_field.send_keys(float(row_data[8].value) if row_data[8].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập Clo Khử trùng.")

            may_1_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Máy 1')]]/input")
            may_1_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            may_1_field.send_keys(int(row_data[9].value) if row_data[9].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập máy 1.")

            may_2_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Máy 2')]]/input")
            may_2_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            may_2_field.send_keys(int(row_data[10].value) if row_data[10].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập máy 2.")

            may_3_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Máy 3')]]/input")
            may_3_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            may_3_field.send_keys(int(row_data[11].value) if row_data[11].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập máy 3.")

            may_4_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Máy 4')]]/input")
            may_4_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            may_4_field.send_keys(int(row_data[12].value) if row_data[12].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập máy 4.")

            may_5_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Máy 5')]]/input")
            may_5_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            may_5_field.send_keys(int(row_data[13].value) if row_data[13].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập máy 5.")

            may_6_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Máy 6')]]/input")
            may_6_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            may_6_field.send_keys(int(row_data[14].value) if row_data[14].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập máy 6.")

            may_7_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Máy 7')]]/input")
            may_7_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            may_7_field.send_keys(int(row_data[15].value) if row_data[15].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập máy 7.")

            may_8_field = driver.find_element(By.XPATH, "//div[label[contains(text(), 'Máy 8')]]/input")
            may_8_field.clear()
            time.sleep(DELAY_UPDATE_INPUT)
            may_8_field.send_keys(int(row_data[16].value) if row_data[16].value is not None else "")
            time.sleep(DELAY_UPDATE_INPUT)
            print("Đã nhập máy 8.")

            # luu_btn = driver.find_element(By.XPATH, "//button[.='Lưu']")
            # luu_btn.click()
            
            print("Hoàn thành tác vụ.")
            driver.quit()
        except Exception as e:
            driver.quit()
            raise e

        return True
    except Exception as e:
        print('Có lỗi xảy ra:', e)
        return False

class AutoQLTBApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.processed = False
        self.current_hour = get_current_hour()
        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.run_task)
        self.xlsx_folder = ""

    def initUI(self):
        self.setWindowTitle('Auto QLTB')
        self.setGeometry(100, 100, 400, 200)
        
        self.start_button = QtWidgets.QPushButton('Start', self)
        self.start_button.setGeometry(150, 80, 100, 30)
        self.start_button.clicked.connect(self.toggle_task)
        
        self.select_folder_button = QtWidgets.QPushButton('Select Folder', self)
        self.select_folder_button.setGeometry(150, 40, 100, 30)
        self.select_folder_button.clicked.connect(self.select_folder)
        
        self.show()

    def select_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.xlsx_folder = folder
            print(f"Selected folder: {self.xlsx_folder}")

    def toggle_task(self):
        if self.timer.isActive():
            self.timer.stop()
            self.start_button.setText('Start')
        else:
            self.timer.start(1000)  # Run every 1 second
            self.start_button.setText('Stop')

    def run_task(self):
        print(".")
        if not self.processed:
            current_minute = get_current_minute()
            result = auto_login_and_input(self.current_hour, current_minute, self.xlsx_folder)
            if result:
                self.processed = True
        
        if self.current_hour != get_current_hour():
            print("Đã qua giờ, cập nhật giờ mới.")
            self.current_hour = get_current_hour()
            self.processed = False

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ex = AutoQLTBApp()
    sys.exit(app.exec_())