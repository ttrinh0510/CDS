import glob
import os
import shutil
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import win32com.client
import pandas as pd
from datetime import datetime, timedelta
import os
# Đặt lại biến môi trường tạm
os.environ['TEMP'] = 'E:\\Temp'
os.environ['TMP'] = 'E:\\Temp'

# Đảm bảo thư mục tồn tại
os.makedirs("E:\\Temp", exist_ok=True)
# Get today's date in the format yyyy-mm-dd
today = datetime.today().strftime('%Y-%m-%d')
error_list = []
year_list = []
checking_date_list = []
# date_of_checking = today
def save_attachments(subject_keyword, save_folder):
    # Tạo đối tượng Outlook
    # error_list.append("- Save attached file from outlook")
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Truy cập hộp thư đến
    mailbox = namespace.Folders["tttrinh1"]
    variety_folder = mailbox.Folders["variety"]
    rpa_folder = variety_folder.Folders["RPA"]
    messages = rpa_folder.Items  # Lấy tất cả các email trong thư mục "RPA"
    messages.Sort("[ReceivedTime]", True)  # Sắp xếp theo thời gian nhận (mới nhất trước)

    # inbox = namespace.GetDefaultFolder(6)  # 6 is the index for the inbox
    # messages = inbox.Items
    # messages.Sort("[ReceivedTime]", True)



    # Tạo thư mục nếu chưa có
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)

    print("Đang tìm kiếm các email có chủ đề: ", subject_keyword)

    # Lặp qua các email
    found = False  # Biến kiểm tra xem có email nào phù hợp không
    for message in messages:
        try:
            # Kiểm tra nếu chủ đề email chứa từ khóa
            if subject_keyword.lower() in message.Subject.lower():
                found = True
                print(f"Đã tìm thấy email với chủ đề: {message.Subject}\n")
                # error_list.append(f"Đã tìm thấy email với chủ đề: {message.Subject}")
                # Lặp qua các tệp đính kèm trong email
                for attachment in message.Attachments:
                    # Định nghĩa đường dẫn để lưu tệp đính kèm
                    if attachment.FileName.lower().endswith(('.xls', '.xlsx')):

                        attachment_file_path = os.path.join(save_folder, attachment.FileName)

                    # Lưu tệp đính kèm
                        attachment.SaveAsFile(attachment_file_path)
                        print(f"Tệp đính kèm đã được lưu: {attachment_file_path}\n")
                        # error_list.append(f"Tệp đính kèm đã được lưu: {attachment_file_path}\n")

        except Exception as e:
            print(f"Lỗi khi xử lý email: {e}")
            error_list.append(f"Lỗi khi xử lý email RPA: {e}")

    if not found:
        print("Không tìm thấy email nào với chủ đề đã cho.")
        error_list.append(f"Không tìm thấy email nào với chủ đề đã cho.")

def extractfile(name):
    # global sheet_name
    # sheet_name = "GENERAL"

    # Construct the file path
    print("- Running extract file\n")
    print(name)
    # error_list.append("--------------------------------Running extract file---------------------------------")
    folder_path = r'\\10.20.254.8\Data\Dept\Admin-Shipping\0-部門公共檔案 SH Public\CDS\RPA DATA'
    filename = f'{name}'
    file_path = os.path.join(folder_path, filename)
    for ext in ['.xlsx', '.xls']:
        file = file_path + ext
        if "EXPORT" in filename or "IMPORT" in filename:
            if os.path.exists(file):
                general_sheet(file, filename)
                detail_sheet(file, filename)
                break

        elif "AMA" in filename:
            if os.path.exists(file):
                ama_report(file, filename)
                break
    else:
        print(f"No file found for {name} with .xlsx or .xls extension.")
        error_list.append(f"No file found for {name} with .xlsx or .xls extension.")
    # Check if the file exists and read it
def general_sheet(file, filename):
    global date_of_checking
    # error_list.append(f"- Running extract general sheet file {file}\n")
    df = pd.read_excel(file,sheet_name="GENERAL", header=None)
    df = df.dropna(how='all')
    df.reset_index(drop=True, inplace=True)
    df.columns = df.iloc[0]
    df = df.drop(index=0)  # Xóa dòng đầu tiên đã được dùng làm header

    import_column_mapping = {
        'STT': 'Update Date',
        'Số TK': 'CDS Number',
        'Số tờ khai đầu tiên': 'First CDS Number',
        'Nhánh': 'Next CDS Number',
        'Ngày ĐK': 'CDS Date',
        'Mã HQ': 'Customs Code',
        'Mã loại hình': 'CDS Type',
        'Số tờ khai tạm nhập tái xuất tương ứng': 'Temporary Import Reexport CDS Number',
        'Tên đối tác': 'Shipper',
        'Mã đại lý hải quan': 'Brokers Code',
        'Bộ phận xử lý': 'Branch',
        'Phương thức vận chuyển': 'Transportation',
        'Vận đơn': 'B/L',
        'Nước nhập khẩu': 'Import Country',
        'Ngày vận đơn': 'B/L Date',
        'Số lượng kiện': 'Package Qty',
        'Tổng trọng lượng hàng (Gross)': 'Gross Weight',
        'Địa điểm lưu kho': 'Storage Location Code',
        'Tên địa điểm lưu kho': 'Storage Location',
        'Mã phương tiện vận chuyển': 'Vessel Code',
        'Tênphương tiện vận chuyển': 'Vessel',
        'Ngày đến': 'ETA',
        'Mã địa điểm dỡ hàng': 'POD Code',
        'Tên địa điểm dỡ hàng': 'POD',
        'Mã địa điểm xếp hàng': 'POL Code',
        'Tên địa điểm xếp hàng': 'POL',
        'Số lượng cont': 'CTN Volume',
        'Số giấy phép': 'Import License',
        'Mã giấy phép': 'Import License Code',
        'Số hóa đơn TM': 'Invoice',
        'Ngày HĐTM': 'Invoice Date',
        'Tổng trị giá hóa đơn': 'Invoice Value (F)',
        'Phương thức thanh toán': 'Payment Term',
        'Điều kiện giá hóa đơn': 'INCOTERM',
        'Ng.Tệ hóa đơn': 'Invoice Currency',
        'Tỷ giá VNĐ': 'Exchange Rate (VND)',
        'Phí BH': 'Insurance',
        'Phí VC': 'Freight',
        'Người nộp thuế': 'Tax Payer',
        'Trị giá KB': 'Declaration Value (F)',
        'Tổng trị giá TT': 'Levied Invoice Value (VND)',
        'Tổng tiền thuế': 'Total Tax',
        'Khoản điều chỉnh 1': 'Adjustment 1',
        'Khoản điều chỉnh 2': 'Adjustment 2',
        'Khoản điều chỉnh 3': 'Adjustment 3',
        'Khoản điều chỉnh 4': 'Adjustment 4',
        'Khoản điều chỉnh 5': 'Adjustment 5',
        'Lý do đề nghị BP': 'Reason for Bank Guarantee Request',
        'Mã ngân hàng trả thuế': 'Bank Code',
        'Tên ngân hàng trả thuế': 'Bank Name',
        'Năm phát hành hạn mức': 'Year of Quota Issuance',
        'Ký hiệu CT hạn mức': 'Quota Certificate Symbol',
        'Số CT hạn mức': 'Quota Certificate Number',
        'Mã xđ thời hạn nộp thuế': 'Tax Payment Term Code',
        'Mã ngân hàng bảo lãnh': 'Bank Guarantee Code',
        'Tên ngân hàng bảo lãnh': 'Bank Guarantee Name',
        'Năm phát hành bảo lãnh': 'Bank Guarantee Year',
        'Ký hiệu CT bảo lãnh': 'Bank Guarantee Symbol',
        'Số hiệu CT bảo lãnh': 'Bank Guarantee Number',
        'Số hợp đồng': 'Contract Number',
        'Ngày hợp đồng': 'Contract Date',
        'Ngày HHHĐ': 'Contract Expiry Date',
        'Trạng thái': 'TQ Status',
        'Phân luồng': 'CDS Status',
        'Ngày thông quan': 'TQ Date',
        'Tổng thuế XNK': 'Tariff Amount',
        'Tổng Thuế TV': 'Safeguard Amount',
        'Số tiền miễn thuế XNK': 'Exemption Amount',
        'Tổng thuế PBĐX': 'Export Tax',
        'Tổng thuế TTĐB': 'Special Tax',
        'Tổng thuế MT': 'Environmental Tax',
        'Tổng thuế VAT': 'Total VAT',
        'Tổng lượng hàng': 'Total Quantity of Goods (NW)',
        'Số hồ sơ TK': 'ROCS',
        'Ghi chú': 'Remark',
        'Số điện thoại': 'Phone',
        'Mã doanh nghiệp': 'Tax Code',
        'Tên doanh nghiệp': 'Company',
        'Mã phân loại khai trị giá': 'Value Declaration Classification Code',
        'Số hợp đồng xuất khẩu': 'Export Contract Number',
        'Số Record': 'Record Number',
        'Nhà cung cấp': 'Partner',
        'Số chứng từ thanh toán': 'Payment Document Number',
        'Ngày chứng từ thanh toán': 'Payment Document Date',
        'Mã tiền tệ phí bảo hiểm': 'Insurance Currency',
        'Mã tiền tệ phí vận chuyển': 'Freight Currency',
        'Địa chỉ đối tác 1': 'Address 1',
        'Địa chỉ đối tác 2': 'Address 2',
        'Địa chỉ đối tác 3': 'Address 3',
        'Địa chỉ đối tác 4': 'Address 4',
        'Ký hiệu và số bao bì': 'Shipping Mark',
        'Declaration Remark': 'Declaration Remark',
        'Giờ thông quan': 'TQ Time',
        'Mã văn bản pháp quy': 'Regulation 1',
        'Mã văn bản pháp quy 2': 'Regulation 2',
        'Mã văn bản pháp quy 3': 'Regulation 3',
        'Mã văn bản pháp quy 4': 'Regulation 4',
        'Mã văn bản pháp quy 5': 'Regulation 5',
        'Mã giấy phép 2': 'License 1',
        'Số giấy phép 2': 'License 2',
        'Mã giấy phép 3': 'License 3',
        'Số giấy phép 3': 'License 4',
        'Mã giấy phép 4': 'License 5',
        'Số giấy phép 4': 'License 6',
        'Mã giấy phép 5': 'License 7',
        'Số giấy phép 5': 'License 8',
        'Chi tiết khai trị giá': 'Declaration Details'
    }

    export_column_mapping = {
        'STT': 'Update Date',
     'Số TK': 'CDS Number',
     'Số tờ khai đầu tiên': 'First CDS Number',
     'Nhánh': 'Branch',
     'Ngày ĐK': 'CDS Date',
     'Mã HQ': 'Customs Code',
     'Mã loại hình': 'CDS Type',
     'Số tờ khai tạm nhập tái xuất tương ứng': 'Temp Import and Reexport CDS Number',
     'Tên đối tác': 'Buyer',
     'Mã đại lý hải quan': 'Customs Broker Code',
     'Bộ phận xử lý': 'Processing Customs',
     'Phương thức vận chuyển': 'Transportation',
     'Vận đơn': 'System Code',
     'Nước nhập khẩu': 'Import Country (F)',
     'Ngày vận đơn': 'B/L Date',
     'Số lượng kiện': 'Package Qty',
     'Tổng trọng lượng hàng (Gross)': 'Gross Weight',
     'Địa điểm lưu kho': 'Storage Location Code',
     'Tên địa điểm lưu kho': 'Storage Location',
     'Mã phương tiện vận chuyển': 'Vessel Code',
     'Tênphương tiện vận chuyển': 'Vessel',
     'Ngày đến': 'ETD',
     'Mã địa điểm xếp hàng': 'POL Code',
     'Tên địa điểm xếp hàng': 'POL',
     'Mã địa điểm dỡ hàng': 'POD Code',
     'Tên địa điểm dỡ hàng': 'POD',
     'Số lượng cont': 'CTN Volume',
     'Số giấy phép': 'Export License Number',
     'Mã giấy phép': 'Export License Code',
     'Số hóa đơn TM': 'Export Invoice',
     'Ngày HĐTM': 'Export Contract Date',
     'Tổng trị giá hóa đơn': 'Export Invoice Value (F)',
     'Phương thức thanh toán': 'Payment Method',
     'Điều kiện giá hóa đơn': 'INCOTERM',
     'Ng.Tệ hóa đơn': 'Invoice Currency',
     'Tỷ giá VNĐ': 'Exchange Rate (VND)',
     'Phí BH': 'Insurance Fee',
     'Phí VC': 'Freight Fee',
     'Người nộp thuế': 'Tax Pyaer Code',
     'Trị giá KB': 'Declaration Value (F)',
     'Tổng trị giá TT': 'Levied Invoice Value (VND)',
     'Tổng tiền thuế': 'Total Tax',
     'Khoản điều chỉnh 1': 'Adjustment 1',
     'Khoản điều chỉnh 2': 'Adjustment 2',
     'Khoản điều chỉnh 3': 'Adjustment 3',
     'Khoản điều chỉnh 4': 'Adjustment 4',
     'Khoản điều chỉnh 5': 'Adjustment 5',
     'Lý do đề nghị BP': 'BP Tax',
     'Mã ngân hàng trả thuế': 'Tax Payment Code',
     'Tên ngân hàng trả thuế': 'Tax Payment Bank',
     'Năm phát hành hạn mức': 'Bank Credit Year',
     'Ký hiệu CT hạn mức': 'Bank Credit Serial Number',
     'Số CT hạn mức': 'Bank Credit Number',
     'Mã xđ thời hạn nộp thuế': 'Tax Payment Type',
     'Mã ngân hàng bảo lãnh': 'Bank Guarantee Code',
     'Tên ngân hàng bảo lãnh': 'Bank Guarantee Name',
     'Năm phát hành bảo lãnh': 'Bank Guarantee Issued Date',
     'Ký hiệu CT bảo lãnh': 'Bank Guarantee Symbol',
     'Số hiệu CT bảo lãnh': 'Bank Guarantee Contract',
     'Số hợp đồng': 'Contract Number',
     'Ngày hợp đồng': 'Contract Date',
     'Ngày HHHĐ': 'Contract Expiry Date',
     'Trạng thái': 'TQ Status',
     'Phân luồng': 'CDS Status',
     'Ngày thông quan': 'TQ Date',
     'Tổng thuế XNK': 'Tariff Amount',
     'Tổng Thuế TV': 'Safeguard Amount',
     'Số tiền miễn thuế XNK': 'Exemption Amount',
     'Tổng thuế PBĐX': 'ADD Amount',
     'Tổng thuế TTĐB': 'Special Consumption Amount',
     'Tổng thuế MT': 'Environmental Tax',
     'Tổng thuế VAT': 'Total VAT',
     'Tổng lượng hàng': 'Regulated Net Weight 1',
     'Số hồ sơ TK': 'On Spot_Customs declaration file number',
     'Ghi chú': 'Remark',
     'Số điện thoại': 'Phone',
     'Mã doanh nghiệp': 'Tax Code',
     'Tên doanh nghiệp': 'Company',
     'Mã phân loại khai trị giá': 'Value declaration classification code',
     'Số hợp đồng xuất khẩu': 'Export Contract Number',
     'Số Record': 'Record Number',
     'Nhà cung cấp': 'Supplier',
     'Số chứng từ thanh toán': 'Payment Document Number',
     'Ngày chứng từ thanh toán': 'Payment Document Date',
     'Mã tiền tệ phí bảo hiểm': 'Insurance Currency',
     'Mã tiền tệ phí vận chuyển': 'Freight Currency',
     'Địa chỉ đối tác 1': 'CNEE Address 1',
     'Địa chỉ đối tác 2': 'CNEE Address 2',
     'Địa chỉ đối tác 3': 'CNEE Address 3',
     'Địa chỉ đối tác 4': 'CNEE Address 4',
     'Ký hiệu và số bao bì': 'Shipping Mark',
     'Chi tiết khai trị giá': 'Value declaration details',
     'Giờ thông quan': 'TQ Time',
     'Mã văn bản pháp quy': 'Regulatory Document Code',
     'Mã văn bản pháp quy 2': 'Regulatory Document Code 2',
     'Mã văn bản pháp quy 3': 'Regulatory Document Code 3',
     'Mã văn bản pháp quy 4': 'Regulatory Document Code 4',
     'Mã văn bản pháp quy 5': 'Regulatory Document Code 5',
     'Mã giấy phép 2': 'License Code 2',
     'Số giấy phép 2': 'License Code 3',
     'Mã giấy phép 3': 'License Code 4',
     'Số giấy phép 3': 'License Code 5',
     'Mã giấy phép 4': 'License Code 6',
     'Số giấy phép 4': 'License Code 7',
     'Mã giấy phép 5': 'License Code 8',
     'Số giấy phép 5': 'License Code 9'
    }

    trans_mapping = {
        'Đường biển (không container)': 'CFS',
        'Đường biển (container)': 'Container',
        'Loại khác': 'Others (VN)',
        'Đường không': 'Air',
        'Đường bộ (xe tải)': 'Overseas Truck'
    }


    TQ_mapping = {
        'Đã phân luồng': 'Pending',
        'Đã thông quan': 'TQ',
        'Nhập mới': 'New'
    }

    CDS_status_mapping = {
        'Luồng vàng': 'Y',
        'Luồng đỏ': 'R',
        'Luồng Xanh': 'G'
    }
    # Đổi tên các cột theo ánh xạ
    if "IMPORT" in file:
        df = df.rename(columns=import_column_mapping)
        df.loc[:, 'Invoice'] = df['Invoice'].astype(str)

    elif "EXPORT" in file:
        df = df.rename(columns=export_column_mapping)
        df.loc[:, 'Export Invoice'] = df['Export Invoice'].astype(str)

    df['Update Date'] = date_of_checking
    df['Transportation'] = df['Transportation'].map(trans_mapping).fillna(df['Transportation'])
    df['TQ Status'] = df['TQ Status'].map(TQ_mapping).fillna(df['TQ Status'])
    df['CDS Status'] = df['CDS Status'].map(CDS_status_mapping).fillna(df['CDS Status'])
    df.loc[:, 'Remark'] = df['Remark'].astype(str)
    df.loc[:, 'Shipping Mark'] = df['Shipping Mark'].astype(str)

    df['Original CDS'] = df['CDS Number'].apply(lambda x: str(x)[:-1])
    df.insert(2, 'Original CDS', df.pop('Original CDS'))

    df['Revision'] = df['CDS Number'].apply(lambda x: str(x)[-1])
    df.insert(3, 'Revision', df.pop('Revision'))

    if "IMPORT" in file:
        df['New B/L'] = df['B/L'].str[6:]
    if "EXPORT" in file:

        for number, inv in enumerate(df['Export Invoice'],start=1):
            # Kiểm tra nếu "Invoice" bắt đầu với "TP" hoặc "FP"
            if pd.notna(inv) and "TP" not in inv and "FP" not in inv:
                remark = df['Remark'][number] # Lấy giá trị của Remark tại dòng hiện tại
                inv = str(inv)
                # print(inv)
                # print(remark)
                # Kiểm tra nếu 'Remark' có chứa "TP" hoặc "FP"
                try:
                    if "TP" in remark:
                        index = remark.find("TP")
                        invoice_number = remark[index:index + 12]  # Lấy 12 ký tự bắt đầu từ "TP"
                        df.loc[number, 'Export Invoice'] = invoice_number
                    elif "FP" in remark:
                        index = remark.find("FP")
                        invoice_number = remark[index:index + 12]  # Lấy 12 ký tự bắt đầu từ "FP"
                        df.loc[number, 'Export Invoice'] = invoice_number
                    elif "TP" in df['Shipping Mark'][number]:
                        index = df['Shipping Mark'][number].find("TP")
                        invoice_number = remark[index:index + 12]  # Lấy 12 ký tự bắt đầu từ "TP"
                        df.loc[number, 'Export Invoice'] = invoice_number
                    elif "FP" in df['Shipping Mark'][number]:
                        index = df['Shipping Mark'][number].find("FP")
                        invoice_number = remark[index:index + 12]  # Lấy 12 ký tự bắt đầu từ "TP"
                        df.loc[number, 'Export Invoice'] = invoice_number


                except Exception as e:
                    if "TP" not in df.loc[number, 'Export Invoice'] or "FP" not in df.loc[number, 'Export Invoice']:
                        error_list.append(f"{inv} is wrong:{e}")
                        print(f"{inv} is wrong in {file}"
                                          f":{e}")


    df.loc[:, 'CDS Number'] = df['CDS Number'].astype(str)
    df['CDS Date'] = pd.to_datetime(df['CDS Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    if "IMPORT" in file:
        df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['ETA'] = pd.to_datetime(df['ETA'], errors='coerce').dt.strftime('%Y-%m-%d')

    elif "EXPORT" in file:
        df['Export Contract Date'] = pd.to_datetime(df['Export Contract Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['ETD'] = pd.to_datetime(df['ETD'], errors='coerce').dt.strftime('%Y-%m-%d')

    df['TQ Date'] = pd.to_datetime(df['TQ Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['B/L Date'] = pd.to_datetime(df['B/L Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['Contract Date'] = pd.to_datetime(df['Contract Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    unique_years = df['CDS Date'].str[:4].unique()  # Lấy 4 ký tự đầu tiên của 'CDS Date' để làm năm
    # df = df.drop_duplicates(subset=['Original CDS'], keep='last')
    for year in unique_years:
        year_list.append(year)
        year_df = df[df['CDS Date'].str.startswith(year)]
        if "EXPORT" in filename:
        # Tạo tên file Excel dựa trên năm
            filename = f"CDS EXPORT REPORT {year} general.xlsx"
        elif "IMPORT" in filename:
            filename = f"CDS IMPORT REPORT {year} general.xlsx"

        # Kiểm tra xem file đã tồn tại chưa
        if os.path.exists(filename):
            # Nếu file đã tồn tại, đọc file cũ vào DataFrame
            existing_df = pd.read_excel(filename)
            # existing_df['CDS Number'] = existing_df['CDS Number'].apply(lambda x: str(x).strip())
            # year_df.loc[:, 'CDS Number'] = year_df['CDS Number'].apply(lambda x: str(x).strip())
            existing_df_filled = existing_df.fillna("")
            year_df_filled = year_df.fillna("")

            combined_df = pd.concat([existing_df_filled, year_df_filled], ignore_index=True)
            # Kết hợp dữ liệu mới và dữ liệu cũ, ưu tiên giữ dòng mới
            # combined_df = pd.concat([existing_df, year_df], ignore_index=True)
            combined_df.loc[:, 'Original CDS'] = combined_df['Original CDS'].apply(lambda x: str(x).strip())
            #
            # combined_df = combined_df.reset_index(drop=True)
            # idx_max = df.groupby('Original CDS')['Revision'].idxmax().dropna().astype(int)
            # valid_idx = idx_max[idx_max < len(combined_df)].astype(int)
            # combined_df = combined_df.iloc[valid_idx]
            ###################################################################################
            # # # Loại bỏ các dòng trùng (giữ dòng mới kết hợp vào)
            combined_df = combined_df.drop_duplicates(subset=['Original CDS'],
                                                      keep='last')  # 'last' giữ dòng mới nhất (dòng vừa thêm vào)
        else:
            # Nếu file chưa tồn tại, chỉ sử dụng dữ liệu mới
            combined_df = year_df

        combined_df.to_excel(filename, index=False, sheet_name="GENERAL")
        print(f"File for year {year} saved as {filename}")

def detail_sheet(file, filename):
    sheet_name = "DETAIL"
    global date_of_checking
    # error_list.append(f"- Running extract detail sheet file {file}\n")
    df = pd.read_excel(file,sheet_name="DETAIL", header=None)
    df = df.dropna(how='all')
    df.reset_index(drop=True, inplace=True)
    df.columns = df.iloc[0]
    df = df.drop(index=0)  # Xóa dòng đầu tiên đã được dùng làm header

    import_column_mapping = {
        'STT': 'Update Date',
        'Số TK': 'CDS Number',
        'Ngày ĐK': 'CDS Date',
        'Mã loại hình': 'CDS Type',
        'Mã địa điểm đích': 'POD Location Code',
        'Tên địa điểm đích cho vận chuyển bảo thuế': 'Bonded Warehouse',
        'Địa điểm dỡ hàng': 'POD',
        'Mã hiệu PTVC': 'Transportation Code',
        'Ngày khởi hành vận chuyển': 'ETD',
        'Ký hiệu và số hiệu bao bì': 'Shipping Mark and Quantity',
        'Tỷ giá thanh toán': 'Exchange Rate (VND)',
        'Đơn vị tiền tệ': 'Exchange Currency',
        'Số lượng kiện': 'Package Qty',
        'Mã ĐVT kiện': 'Package Unit',
        'Trọng lượng': 'Gross Weight',
        'Mã ĐVT trọng lượng': 'Gross Weight Unit',
        'Số quản lý nội bộ': 'ROCS',
        'Điều kiện giá hóa đơn': 'INCOTERM',
        'Ghi chú': 'Remark',
        'STT hàng': 'CDS Rows Number',
        'Mã NPL/SP': 'Material Code',
        'Mã HS': 'HS Code',
        'Tên hàng': 'Description of Goods',
        'Xuất xứ': 'Country of Origin',
        'Đơn giá': 'Unit Price (F)',
        'Đơn giá tính thuế': 'Unit Price (VND)',
        'Tổng số lượng': 'Tracking Net Weight 1',
        'Đơn vị tính': 'Tracking Net Weight 1 Unit',
        'Tổng số lượng 2': 'Regulated Net Weight 2',
        'Đơn vị tính 2': 'Regulated Net Weight 2 Unit',
        'Trị giá NT': 'Invoice Value (F)',
        'Tổng trị giá': 'Levied Unit Price (VND)',
        'Mã biểu thuế XNK': 'Favorable Tariff Code',
        'Thuế suất XNK': 'Tariff',
        'Tiền thuế XNK': 'Tariff Amount',
        'Số tiền miễn thuế XNK': 'Exemption Amount',
        'Thuế suất TV': 'Safeguard Rate',
        'Tiền thuế TV': 'Safeguard Amount',
        'Thuế suất PB': 'Export Tax Rate',
        'Tiền thuế PB': 'Export Tax Amount',
        'Thuế suất TTĐB': 'Special Tax',
        'Tiền thuế TTĐB': 'Special Tax Amount',
        'Thuế suất BVMT': 'Environmental Tax',
        'Tiền thuế MT': 'Environmental Tax Amount',
        'Thuế suất VA': 'VAT Rate',
        'Tiền thuế VAT': 'VAT Amount',
        'Tổng tiền thuế': 'Total Tax',
        'Mã doanh nghiệp': 'Tax Code',
        'Tên doanh nghiệp': 'Company',
        'Tên đối tác': 'Shipper',
        'Số hóa đơn': 'Invoice',
        'Ngày hóa đơn': 'Invoice Date',
        'Số hợp đồng': 'PO Number',
        'Ngày hợp đồng': 'Contract Date',
        'Chi tiết khai trị giá': 'Declaration Details',
    }
    export_column_mapping = {
        "STT": "Update Date",
        "Số TK": "CDS Number",
        "Ngày ĐK": "CDS Date",
        "Mã loại hình": "CDS Type",
        "Mã địa điểm đích": "Place of Receipt Code",
        "Tên địa điểm đích cho vận chuyển bảo thuế": "Place of Receipt",
        "Địa điểm dỡ hàng": "POD",
        "Mã hiệu PTVC": "Transportation Code",
        "Ngày khởi hành vận chuyển": "Drop Off Date",
        "Ký hiệu và số hiệu bao bì": "Shipping Mark",
        "Tỷ giá thanh toán": "Exchange Rate (VND)",
        "Đơn vị tiền tệ": "Invoice Currency",
        "Số lượng kiện": "Package Qty",
        "Mã ĐVT kiện": "Package Unit",
        "Trọng lượng": "Gross Weight",
        "Mã ĐVT trọng lượng": "Gross Weight Unit",
        "Số quản lý nội bộ": "On Spot_Customs declaration file number",
        "Điều kiện giá hóa đơn": "INCOTERM",
        "Ghi chú": "Remark",
        "STT hàng": "CDS Rows Number",
        "Mã NPL/SP": "Material Code",
        "Mã HS": "HS Code",
        "Tên hàng": "Description of Goods",
        "Xuất xứ": "Country of Origin",
        "Đơn giá": "Unit Price (F)",
        "Đơn giá tính thuế": "Levied Unit Price (VND)",
        "Tổng số lượng": "Tracking Net Weight 1",
        "Đơn vị tính": "Tracking Net Weight 1 Unit",
        "Tổng số lượng 2": "Regulated Net Weight 2",
        "Đơn vị tính 2": "Regulated Net Weight 2 Unit",
        "Trị giá NT": "Declaration Value (F)",
        "Tổng trị giá": "Levied Invoice Value (VND)",
        "Mã biểu thuế XNK": "Tariff Code",
        "Thuế suất XNK": "Tariff Rate",
        "Tiền thuế XNK": "Tariff Amount",
        "Số tiền miễn thuế XNK": "Exemption Amount",
        "Thuế suất TV": "Special TV Rate",
        "Tiền thuế TV": "Special TV Amount",
        "Thuế suất PB": "ADD Rate",
        "Tiền thuế PB": "ADD Amount",
        "Thuế suất TTĐB": "Special Consumption Rate",
        "Tiền thuế TTĐB": "Special Consumption Amount",
        "Thuế suất BVMT": "Environmental Rate",
        "Tiền thuế MT": "Environmental Amount",
        "Thuế suất VA": "VAT Rate",
        "Tiền thuế VAT": "Vat Amount",
        "Tổng tiền thuế": "Total Tax",
        "Mã doanh nghiệp": "Tax Code",
        "Tên doanh nghiệp": "Company",
        "Tên đối tác": "Buyer",
        "Số hóa đơn": "Invoice",
        "Ngày hóa đơn": "Invoice Date",
        "Số hợp đồng": "Contract Number",
        "Ngày hợp đồng": "Contract Date"
    }

    transportcode_mapping = {
        2: 'Container',
        9: 'Overseas Truck',
        4: 'Others (VN)',
        1: 'Air',
        3: 'CFS',
    }
    # Đổi tên các cột theo ánh xạ
    if "IMPORT" in file:
        df = df.rename(columns=import_column_mapping)
        df.loc[:, 'Invoice'] = df['Invoice'].astype(str)
        df.loc[:, 'Shipping Mark and Quantity'] = df['Shipping Mark and Quantity'].astype(str)

    elif "EXPORT" in file:
        df = df.rename(columns=export_column_mapping)
        df.loc[:, 'Invoice'] = df['Invoice'].astype(str)
        df.loc[:, 'Shipping Mark'] = df['Shipping Mark'].astype(str)

    df.loc[:, 'Remark'] = df['Remark'].astype(str)

    df['Original CDS'] = df['CDS Number'].apply(lambda x: str(x)[:-1])
    df.insert(2, 'Original CDS', df.pop('Original CDS'))

    df['Revision'] = df['CDS Number'].apply(lambda x: str(x)[-1])
    df.insert(3, 'Revision', df.pop('Revision'))
    # df['Transportation Code'] = df['Transportation Code'].map(transportcode_mapping)
    # Ensure the 'Transportation Code' column is converted to string before mapping
    df['Transportation Code'] = df['Transportation Code'].astype(str).map(
        {str(k): v for k, v in transportcode_mapping.items()}
    ).fillna(df['Transportation Code'])

    if "EXPORT" in file:
        for number, inv in enumerate(df['Invoice'],start=1):
            # Kiểm tra nếu "Invoice" bắt đầu với "TP" hoặc "FP"
            inv = str(inv)
            if pd.notna(inv) and (not inv.startswith("TP") and not inv.startswith("FP")):
                remark = df['Remark'][number] # Lấy giá trị của Remark tại dòng hiện tại

                # Kiểm tra nếu 'Remark' có chứa "TP" hoặc "FP"
                try:
                    if "TP" in remark:
                        index = remark.find("TP")
                        invoice_number = remark[index:index + 12]  # Lấy 12 ký tự bắt đầu từ "TP"
                        df.loc[number, 'Invoice'] = invoice_number
                    elif "FP" in remark:
                        index = remark.find("FP")
                        invoice_number = remark[index:index + 12]  # Lấy 12 ký tự bắt đầu từ "FP"
                        df.loc[number, 'Invoice'] = invoice_number
                    elif "TP" in df['Shipping Mark'][number]:
                        index = df['Shipping Mark'][number].find("TP")
                        invoice_number = remark[index:index + 12]  # Lấy 12 ký tự bắt đầu từ "TP"
                        df.loc[number, 'Invoice'] = invoice_number
                    elif "FP" in df['Shipping Mark'][number]:
                        index = df['Shipping Mark'][number].find("FP")
                        invoice_number = remark[index:index + 12]  # Lấy 12 ký tự bắt đầu từ "TP"
                        df.loc[number, 'Invoice'] = invoice_number

                except Exception as e:
                    if "TP" not in df.loc[number, 'Invoice'] or "FP" not in df.loc[number, 'Invoice']:
                        error_list.append(f"{inv} is wrong:{e}")

                        print(f"{inv} is wrong in {file}"
                              f":{e}")
    if "IMPORT" in file:
        df['Update Date'] = date_of_checking
        df.loc[:, 'CDS Number'] = df['CDS Number'].astype(str)
        df['CDS Date'] = pd.to_datetime(df['CDS Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['ETD'] = pd.to_datetime(df['ETD'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['Contract Date'] = pd.to_datetime(df['Contract Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    elif "EXPORT" in file:
        df['Update Date'] = date_of_checking
        df.loc[:, 'CDS Number'] = df['CDS Number'].astype(str)
        df['CDS Date'] = pd.to_datetime(df['CDS Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['Drop Off Date'] = pd.to_datetime(df['Drop Off Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['Contract Date'] = pd.to_datetime(df['Contract Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    # df = df.drop_duplicates(subset=['CDS Number', 'CDS Rows Number'], keep='last')

    unique_years = df['CDS Date'].str[:4].unique()  # Lấy 4 ký tự đầu tiên của 'CDS Date' để làm năm

    for year in unique_years:
        year_list.append(year)
        year_df = df[df['CDS Date'].str.startswith(year)]
        if "EXPORT" in filename:
        # Tạo tên file Excel dựa trên năm
            filename = f"CDS EXPORT REPORT {year} detail.xlsx"
        elif "IMPORT" in filename:
            filename = f"CDS IMPORT REPORT {year} detail.xlsx"

        # Kiểm tra xem file đã tồn tại chưa
        if os.path.exists(filename):
            # Nếu file đã tồn tại, đọc file cũ vào DataFrame
            existing_df = pd.read_excel(filename)
            # existing_df['CDS Number'] = existing_df['CDS Number'].apply(lambda x: str(x).strip())
            # year_df.loc[:, 'CDS Number'] = year_df['CDS Number'].apply(lambda x: str(x).strip())
            # existing_df_cleaned = existing_df.dropna(axis=1, how='all')
            # year_df_cleaned = year_df.dropna(axis=1, how='all')
            # Kết hợp dữ liệu mới và dữ liệu cũ, ưu tiên giữ dòng mới
            existing_df_filled = existing_df.fillna("")
            year_df_filled = year_df.fillna("")

            combined_df = pd.concat([existing_df_filled, year_df_filled], ignore_index=True)
            # combined_df = pd.concat([existing_df, year_df], ignore_index=True)
            if "IMPORT" in filename:
                combined_df.loc[:, 'CDS Number'] = combined_df['CDS Number'].apply(lambda x: str(x).strip())
                combined_df.loc[:, 'Material Code'] = combined_df['Material Code'].apply(lambda x: str(x).strip())
                combined_df['Invoice Value (F)'] = combined_df['Invoice Value (F)'].apply(lambda x: round(float(str(x).strip()), 2))
                combined_df.loc[:, 'CDS Type'] = combined_df['CDS Type'].apply(lambda x: str(x).strip())
                combined_df['HS Code'] = combined_df['HS Code'].apply(lambda x: round(float(str(x).strip()), 0))
                combined_df.loc[:, 'Invoice'] = combined_df['Invoice'].apply(lambda x: str(x).strip())
                combined_df['CDS Rows Number'] = combined_df['CDS Rows Number'].apply(lambda x: round(float(str(x).strip()), 0))
            elif "EXPORT" in filename:
                combined_df.loc[:, 'CDS Number'] = combined_df['CDS Number'].apply(lambda x: str(x).strip())
                combined_df.loc[:, 'Material Code'] = combined_df['Material Code'].apply(lambda x: str(x).strip())
                combined_df['Declaration Value (F)'] = combined_df['Declaration Value (F)'].apply(
                    lambda x: round(float(str(x).strip()), 2))
                combined_df.loc[:, 'CDS Type'] = combined_df['CDS Type'].apply(lambda x: str(x).strip())
                combined_df['HS Code'] = combined_df['HS Code'].apply(lambda x: round(float(str(x).strip()), 0))
                combined_df.loc[:, 'Invoice'] = combined_df['Invoice'].apply(lambda x: str(x).strip())
                combined_df['CDS Rows Number'] = combined_df['CDS Rows Number'].apply(
                    lambda x: round(float(str(x).strip()), 0))
            # Loại bỏ các dòng trùng (giữ dòng mới kết hợp vào)
            combined_df = combined_df.drop_duplicates(subset=['CDS Number', 'CDS Rows Number'],
                                                      keep='last')  # 'last' giữ dòng mới nhất (dòng vừa thêm vào)
            #######################
            # # Đảm bảo index không bị mất hoặc sai lệch
            # combined_df = combined_df.reset_index(drop=True)
            #
            # # Lấy chỉ mục của dòng có 'Update Date' lớn nhất trong từng nhóm
            # idx_max = combined_df.groupby(['CDS Number', 'CDS Rows Number'])['Update Date'].idxmax().dropna().astype(
            #     int)
            #
            # # Chỉ lấy những index hợp lệ để tránh KeyError
            # valid_idx = idx_max[idx_max.isin(combined_df.index)]
            #
            # # Lọc DataFrame với các index hợp lệ
            # combined_df = combined_df.iloc[valid_idx]


        else:
            # Nếu file chưa tồn tại, chỉ sử dụng dữ liệu mới
            combined_df = year_df
        # Kiểm tra xem thư mục đã tồn tại chưa

        # Lưu DataFrame vào file Excel
        print(f"File for year {year} saved as {filename}")
        combined_df.to_excel(filename, index=False, sheet_name="DETAIL")

        # with pd.ExcelWriter(filename, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        #     combined_df.to_excel(writer, index=False, sheet_name="DETAIL")
def combine_file():
    # Định nghĩa thư mục chứa các tệp Excel
    directory = r"C:\Users\tttrinh\PycharmProjects\FormatCDS"  # Thay đổi đường dẫn tới thư mục chứa tệp của bạn

    # Lấy danh sách tất cả các tệp Excel trong thư mục
    excel_files = [f for f in os.listdir(directory) if f.endswith(".xlsx")]

    # Dictionnary lưu trữ các tệp "general" và "detail" theo năm cho mỗi loại
    files_by_year = {"EXPORT": {}, "IMPORT": {}}

    # Phân loại các tệp vào dictionary theo năm và loại (EXPORT, IMPORT)
    for file in excel_files:
        # Lấy năm từ tên tệp (giả sử năm luôn ở giữa "CDS EXPORT REPORT {year}" trong tên tệp)
        if "general" in file or "detail" in file:
            year = file.split(" ")[3]  # Tên tệp có dạng "CDS EXPORT REPORT {year} general.xlsx" hoặc "detail.xlsx"
            file_type = "EXPORT" if "EXPORT" in file else "IMPORT"  # Kiểm tra xem là EXPORT hay IMPORT

            # Thêm tệp vào dictionary theo loại (EXPORT hoặc IMPORT) và năm
            if year not in files_by_year[file_type]:
                files_by_year[file_type][year] = {"general": None, "detail": None}

            if "general" in file:
                files_by_year[file_type][year]["general"] = file
            elif "detail" in file:
                files_by_year[file_type][year]["detail"] = file

    # Kết hợp các tệp và lưu vào tệp Excel mới cho mỗi năm và mỗi loại (EXPORT, IMPORT)
    for file_type in files_by_year:
        for year, files in files_by_year[file_type].items():
            general_filename = os.path.join(directory, files["general"]) if files["general"] else None
            detail_filename = os.path.join(directory, files["detail"]) if files["detail"] else None
            combined_filename = os.path.join(directory, f"CDS {file_type} REPORT {year}.xlsx")

            # Kiểm tra xem cả hai tệp general và detail đều có tồn tại không
            if general_filename and detail_filename:
                try:
                    # Đọc dữ liệu từ tệp general và detail
                    general_df = pd.read_excel(general_filename)
                    detail_df = pd.read_excel(detail_filename)

                    # Ghi vào một tệp Excel với hai sheet GENERAL và DETAIL
                    with pd.ExcelWriter(combined_filename, engine="openpyxl") as writer:
                        general_df.to_excel(writer, index=False, sheet_name="GENERAL")
                        detail_df.to_excel(writer, index=False, sheet_name="DETAIL")

                    print(f"Combined file for {file_type} {year} saved as {combined_filename}")
                    destination_directory = r"\\10.20.254.8\Data\Dept\Admin-Shipping\0-部門公共檔案 SH Public\CDS"
                    shutil.copy(combined_filename, destination_directory)
                    print("copy the file to destination")
                except Exception as e:
                    print(f"Error processing files for {file_type} {year}: {e}")
                    error_list.append(f"Error processing files for {file_type} {year}: {e}")
            else:
                print(f"Missing 'general' or 'detail' file for {file_type} {year}. Skipping.")
                error_list.append(f"Missing 'general' or 'detail' file for {file_type} {year}. Skipping.")

def combine_file2():
    global year_list
    year_list = list(set(year_list))
    print(year_list)
    if len(year_list) >= 1:
        error_list.append(f"- Running combine file with year list: {year_list}\n")

        # Định nghĩa thư mục chứa các tệp Excel
        directory = r"C:\Users\tttrinh\PycharmProjects\FormatCDS"  # Thay đổi đường dẫn tới thư mục chứa tệp của bạn

        # Lấy danh sách tất cả các tệp Excel trong thư mục
        excel_files = [f for f in os.listdir(directory) if f.endswith(".xlsx")]

        # Dictionnary lưu trữ các tệp "general" và "detail" theo năm cho mỗi loại
        files_by_year = {"EXPORT": {}, "IMPORT": {}}

        # Phân loại các tệp vào dictionary theo năm và loại (EXPORT, IMPORT)
        for file in excel_files:
            if "general" in file or "detail" in file:
                # Lấy năm từ tên tệp (giả sử năm luôn ở giữa "CDS EXPORT REPORT {year}" trong tên tệp)
                year = file.split(" ")[3]  # Tên tệp có dạng "CDS EXPORT REPORT {year} general.xlsx" hoặc "detail.xlsx"
                file_type = "EXPORT" if "EXPORT" in file else "IMPORT"  # Kiểm tra xem là EXPORT hay IMPORT

                # Thêm tệp vào dictionary theo loại (EXPORT hoặc IMPORT) và năm
                if year not in files_by_year[file_type]:
                    files_by_year[file_type][year] = {"general": None, "detail": None}

                if "general" in file:
                    files_by_year[file_type][year]["general"] = file
                elif "detail" in file:
                    files_by_year[file_type][year]["detail"] = file

        # Kết hợp các tệp và lưu vào tệp Excel mới cho mỗi năm và mỗi loại (EXPORT, IMPORT)
        for file_type in files_by_year:
            for year, files in files_by_year[file_type].items():
                # Chỉ xử lý các tệp của các năm có trong danh sách năm yêu cầu
                if year not in year_list:
                    continue

                general_filename = os.path.join(directory, files["general"]) if files["general"] else None
                detail_filename = os.path.join(directory, files["detail"]) if files["detail"] else None
                combined_filename = os.path.join(directory, f"CDS {file_type} REPORT {year}.xlsx")

                # Kiểm tra xem cả hai tệp general và detail đều có tồn tại không
                if general_filename and detail_filename:
                    try:
                        # Đọc dữ liệu từ tệp general và detail
                        general_df = pd.read_excel(general_filename)
                        detail_df = pd.read_excel(detail_filename)

                        general_df['CDS Number'] = general_df['CDS Number'].apply(lambda x: str(x).strip())
                        general_df['Original CDS'] = general_df['Original CDS'].apply(lambda x: str(x).strip())

                        detail_df['CDS Number'] = detail_df['CDS Number'].apply(lambda x: str(x).strip())
                        detail_df['Original CDS'] = detail_df['Original CDS'].apply(lambda x: str(x).strip())

                        # Ghi vào một tệp Excel với hai sheet GENERAL và DETAIL
                        with pd.ExcelWriter(combined_filename, engine="openpyxl") as writer:
                            general_df.to_excel(writer, index=False, sheet_name="GENERAL")
                            detail_df.to_excel(writer, index=False, sheet_name="DETAIL")

                        # recheck(f"CDS {file_type} REPORT {year}.xlsx")

                        print(f"Combined file for {file_type} {year} saved as {combined_filename}")
                        if "EXPORT" in combined_filename:
                            destination_directory = r"\\10.20.254.8\Data\Dept\Admin-Shipping\0-部門公共檔案 SH Public\CDS\Export"
                        elif "IMPORT" in combined_filename:
                            destination_directory = r"\\10.20.254.8\Data\Dept\Admin-Shipping\0-部門公共檔案 SH Public\CDS\Import"
                        else:
                            destination_directory = r"\\10.20.254.8\Data\Dept\Admin-Shipping\0-部門公共檔案 SH Public\CDS"

                        shutil.copy(combined_filename, destination_directory)
                        print("Copy the file to destination")
                        error_list.append(f"Success {file_type} {year} saved as {combined_filename} ")
                    except Exception as e:
                        print(f"Error processing files for {file_type} {year}: {e}")
                        error_list.append(f"Error processing files for {file_type} {year}: {e}")
                else:
                    print(f"Missing 'general' or 'detail' file for {file_type} {year}. Skipping.")
                    error_list.append(f"Missing 'general' or 'detail' file for {file_type} {year}. Skipping.")
    else:
        error_list.append("Nothing to combine")
    # Gọi hàm với danh sách các năm mà bạn muốn gộp

def sendmail():
    today = datetime.now()
    today = today.strftime("%Y-%m-%d")
    my_mail = "trackingatd.janice@gmail.com"
    my_pass = "tgpx kvik odcb dedi"

    # directory = r"C:\Users\tttrinh\PycharmProjects\FormatCDS"
    # files_to_send = glob.glob(os.path.join(directory, "*error*.xlsx"))

    msg = MIMEMultipart()
    msg['Subject'] = f"[AUTOMATION] FORMAT CDS"
    msg['From'] = my_mail
    # msg['To'] = "tttrinh@fenc.vn, tgan@fenc.vn, jerryyang@fenc.com, ttrinh0510@gmail.com"
    msg['To'] = "tttrinh@fenc.vn"

    # for file in files_to_send:
    #     part = MIMEBase('application', "octet-stream")
    #     with open(file, "rb") as f:
    #         part.set_payload(f.read())
    #
    #     # Mã hóa file thành base64
    #     encoders.encode_base64(part)
    #     # Đặt tên cho file đính kèm
    #     part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(file)}")
    #     # Thêm file đính kèm vào email
    #     msg.attach(part)

    error_list_str = "\n".join(error_list) if error_list else "None"
    date_str = "\n".join(checking_date_list) if checking_date_list else "None"
    text = (f"Checking date:\n"
            f"{date_str}"
            f"{today}\n"
            f"{error_list_str}\n\n"
            # f"{upload_note}"
            "\n"
            "Running schedule:\n"
            "7:00 AM everyday")
    # Attach the email body
    msg.attach(MIMEText(text))


    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(user=my_mail, password=my_pass)
            server.send_message(msg)
        print("Mail sent successfully.")
    except:
        print(f"Failed to send email")

def recheck(file):
    print(f"Rechecking {file}")
    general_df = pd.read_excel(file, sheet_name="GENERAL")
    detail_df = pd.read_excel(file, sheet_name="DETAIL")
    file_name = file.replace(".xlsx", "")

    # Làm tròn giá trị trong cột 'Invoice Value' của cả general_df và detail_df (2 chữ số thập phân)
    general_df['Invoice Value'] = general_df['Invoice Value'].round(2)
    detail_df['Invoice Value'] = detail_df['Invoice Value'].round(2)

    # Nhóm theo 'CDS Number' và tính tổng 'Invoice Value' trong detail_df
    total_invoice_value_detail = detail_df.groupby('CDS Number')['Invoice Value'].sum().reset_index()

    # Kết nối dữ liệu của general_df và total_invoice_value_detail dựa trên cột 'CDS Number'
    merged_df = pd.merge(general_df, total_invoice_value_detail, on='CDS Number', how='left',
                         suffixes=('_general', '_detail'))

    # Làm tròn giá trị của cột 'Invoice Value' sau khi merge (để so sánh)
    merged_df['Invoice Value_general'] = merged_df['Invoice Value_general'].round(0)
    merged_df['Invoice Value_detail'] = merged_df['Invoice Value_detail'].round(0)
    merged_df['Declaration Value'] = merged_df['Declaration Value'].round(0)
    # Kiểm tra nếu tổng 'Invoice Value' trong detail_df khớp với 'Invoice Value' trong general_df
    merged_df['is_valid'] = merged_df['Declaration Value'] == merged_df['Invoice Value_detail']

    # Lọc ra những dòng không hợp lệ
    invalid_rows = merged_df[merged_df['is_valid'] == False]

    # Nếu có dòng không hợp lệ, xuất ra file Excel
    if not invalid_rows.empty:
        columns_to_export = ['CDS Date', 'CDS Number', 'Invoice', 'Invoice Value_general','Declaration Value', 'Invoice Value_detail']

        invalid_rows[columns_to_export].to_excel(f"error {file_name}.xlsx", index=False)

        print(f"Report saved as report {file_name}.xlsx")
    else:
        print("Tất cả các invoice đều hợp lệ.")
        error_list.append(f"{file_name} is correct\n")

def remove_duplicate():
    folder_path = r"C:\Users\tttrinh\PycharmProjects\FormatCDS"

    # Tìm tất cả các file .xlsx có chữ "GENERAL" trong tên
    file_list = glob.glob(os.path.join(folder_path, "*general*.xlsx"))

    # Xử lý từng file
    for file in file_list:
        # Đọc file Excel
        df = pd.read_excel(file)

        # Giữ lại giá trị lớn nhất trong cột 'Revision' cho mỗi 'Original CDS'
        df_grouped = df.loc[df.groupby('Original CDS')['Revision'].idxmax()]

        # Ghi đè lên chính file đó
        df_grouped.to_excel(file, index=False, sheet_name="GENERAL")

        print(f"Đã xử lý và lưu lại: {file}")

    print("Hoàn thành!")



def revise_column_name():
    import pandas as pd
    import os

    # Define the folder path
    folder_path = r"C:\Users\tttrinh\PycharmProjects\FormatCDS"

    # Define column mappings for "general" files
    general_column_mapping = {
        "Import License number": "Import License Code",
        "Invoice Value": "Invoice Value (F)",
        "Declaration Value": "Declaration Value (F)",
        "Invoice (VND)": "Levied Invoice Value (VND)",
        "Total Quantity of Goods": "Total Quantity of Goods (NW)",
    }

    # Define column mappings for "detail" files
    detail_column_mapping = {
        "Unit Price": "Unit Price (F)",
        "Unit Price (VND)": "Levied Unit Price (VND)",
        "Invoice Value": "Invoice Value (F)",
        "Invoice Value (VND)": "Levied Invoice Value (VND)",
    }

    # Process all Excel files in the folder
    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)

        if file.endswith(".xlsx"):  # Ensure only Excel files are processed
            if "general" in file.lower():
                sheet_name = "GENERAL"
                column_mapping = general_column_mapping
            elif "detail" in file.lower():
                sheet_name = "DETAIL"
                column_mapping = detail_column_mapping
            else:
                continue  # Skip files that don't match criteria

            try:
                # Read the Excel file with the specified sheet
                df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
                df['Contract Date'] = pd.to_datetime(df['Contract Date'], errors='coerce').dt.strftime('%Y-%m-%d')
                df['CDS Date'] = pd.to_datetime(df['CDS Date'], errors='coerce').dt.strftime('%Y-%m-%d')
                df['Contract Date'] = pd.to_datetime(df['Contract Date'], errors='coerce').dt.strftime('%Y-%m-%d')
                # Rename columns
                df.rename(columns=column_mapping, inplace=True)

                # Save back to Excel, ensuring the sheet name remains unchanged
                with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                print(f"Updated column names in {file} (Sheet: {sheet_name})")

            except Exception as e:
                print(f"Error processing {file}: {e}")




def format_file(file_path, sheet_name):

    # Read the Excel file with the specified sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    df['Contract Date'] = pd.to_datetime(df['Contract Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['CDS Date'] = pd.to_datetime(df['CDS Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['Contract Date'] = pd.to_datetime(df['Contract Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    # Rename columns

    # Save back to Excel, ensuring the sheet name remains unchanged
    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

def rename_duplicate_columns(columns):
    seen = {}
    new_columns = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            new_columns.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_columns.append(col)
    return new_columns
def ama_report(file, filename):
    global date_of_checking
    # error_list.append(f"- Running extract general sheet file {file}\n")
    df = pd.read_excel(file, sheet_name="AMA", header=None)
    df = df.dropna(how='all')
    df.reset_index(drop=True, inplace=True)
    df.columns = df.iloc[0]
    df = df.drop(index=0)  # Xóa dòng đầu tiên đã được dùng làm header

    # print(df)
    # Define the mapping of Vietnamese column names to English
    # Creating a dictionary with Vietnamese column names as keys and English column names as values
    ama_columns_mapping = {
        "STT": "Update Date",
        "Số tờ khai bổ sung": "Amendament Number",
        "Ngày đăng ký": "Amendament Date",
        "Giờ đăng ký": "Amendament Time",
        "Mã Hải quan": "Customs Code",
        "Tên hải quan": "Customs Branch Code",
        "Nhóm xử lý": "Processing",
        "Phân loại xuất nhập khẩu": "Classification of Import and Export",
        "Số tờ khai ban đầu": "Original CDS",
        "Mã loại hình": "CDS Type",
        "Ngày khai báo nhập xuất": "Original CDS Date",
        "Ngày cấp phép nhập xuất": "Approval Date",
        "Thời hạn tạm nhập/tái xuất": "Temporary Import/Re-export Duration",
        "Mã số thuế người khai": "Tax Code",
        "Tên người khai báo": "Company",
        "Mã bưu chính": "Zip Code",
        "Địa chỉ người khai": "Address",
        "Số điện thoại": "Phone",
        "Mã đại lý": "Customs Broker Code",
        "Mã lý do bổ sung": "Value Declaration Classification Code",
        "Mã loại nộp thuế": "Tax Payment Code",
        "Mã Ng.Tệ tiền thuế": "Tax Currency",
        "Mã ngân hàng trả thuế thay": "Bank Code",
        "Năm phát hành": "Issued Year",
        "Ký hiệu chứng từ": "Document Code",
        "Số chứng từ": "Document Number",
        "Ngày hết hạn": "Expiry Date",
        "Ân hạn số ngày": "Grace Periods",
        "Mã ngân hàng bảo lãnh": "Bank Guarantee Code",
        "Năm bảo lãnh": "Guarantee Year",
        "Ký hiệu bảo lãnh": "Guarantee Code",
        "Bảo lãnh số chứng từ": "Guarantee Document",
        "Mã ng.tệ trước khi khai bổ sung": "Currency Code",
        "Tỷ giá ng.tệ trước khi khai bổ sung": "Exchange Rate",
        "Mã ng.tệ sau khi khai bổ sung": "Currency After",
        "Tỷ giá ng.tệ sau khi khai bổ sung": "Exchange Rate After",
        "Số quản lý nội bộ": "Invoice Number",
        "Ghi chú trước khi khai bổ sung": "Notes Before",
        "Số quản lý nội bộ_1": "Invoice Number 1",
        "Ghi chú sau khi khai bổ sung": "Notes After",
        "Ghi chú trước khi khai bổ sung_1": "Reason",
        "Số thông báo": "Notification Number",
        "Ngày hoàn thành kiểm tra": "Inspection Complete Date",
        "Giờ hoàn thành kiểm tra": "Inspection Complete Time",
        "Số trang tờ khai": "Declaration Pages",
        "Số dòng hàng": "Items Lines",
        "Lý do bổ sung": "Supplementary Reason",
        "Người phụ trách duyệt của HQ": "HQ Approval Officer",
        "Tên chi cục HQ tiếp nhận": "Receiving Customs Branch",
        "Ngày đăng ký dữ liệu": "Data Registration Date",
        "Giờ đăng ký dữ liệu": "Data Registration Time",
        "Số thứ tự hàng": "Serial Number",
        "Mã NPL,SP trước khi sửa": "Material Code Before",
        "Tên hàng trước khi sửa": "Description of Goods Before",
        "Mã NPL,SP  sau khi sửa": "Material Code After",
        "Tên hàng sau khi sửa": "Description of Goods After",
        "Nước xuất xứ trước khi sửa": "Country of Origin Before",
        "Nước xuất xứ sau khi sửa": "Country of Origin After",
        "Trị giá tính thuế trước khi sửa": "Levied Invoice Value Before (VND)",
        "Lượng trước khi sửa": "Quantity Before",
        "Mã đơn vị tính trước khi sửa": "Unit Before",
        "Mã HS trước khi sửa": "HS Code Before",
        "Thuế xuất nhập khẩu trước khi sửa": "Tariff Rate Before",
        "Tiền thuế trước khi sửa": "Tax Amount Before",
        "Miễn thuế trước khi sửa": "Tax Exemption Before",
        "Trị giá tính thuế sau khi sửa": "Levied Invoice Value After (VND)",
        "Lượng sau khi sửa": "Quantity After",
        "Mã đơn vị tính sau khi sửa": "Unit After",
        "Mã HS sau khi sửa": "HS Code After",
        "Thuế xuất nhập khẩu sau khi sửa": "Tariff Rate After",
        "Tiền thuế sau khi sửa": "Tax Amount After",
        "Miễn thuế sau khi sửa": "Tax Exemption After",
        "Số tiền thuế điều chỉnh": "Adjusted Tax Amount",
        "Số tiền điều chỉnh": "Adjusted Amount",
        "Mã nguyên tệ điều chỉnh": "Adjusted Currency",
        "Miễn thuế khác 4 sau khi sửa_1": "Other Tax Exemption Before 4_1",
        "Số tiền điều chỉnh thuế khác 4_1": "Adjusted Other Tax Amount 4_1",
    }

    # Extending for "Other Taxes"
    for i in range(1, 75):
        ama_columns_mapping[f"Mã thuế khác {i}"] = f"Other Tax Code {i}"
        ama_columns_mapping[f"Tên thuế khác {i}"] = f"Other Tax {i}"
        ama_columns_mapping[f"Trị giá TT khác {i} trước khi sửa"] = f"Other Tax Value Before {i}"
        ama_columns_mapping[f"Lượng TT khác {i} trước khi sửa"] = f"Other Quantity Before {i}"
        ama_columns_mapping[f"Mã DVT thuế khác {i} trước khi sửa"] = f"Other Unit Before {i}"
        ama_columns_mapping[f"Mã áp dụng thuế khác {i} trước khi sửa"] = f"Other Tax Code Before {i}"
        ama_columns_mapping[f"Thuế suất khác {i} trước khi sửa"] = f"Other Tax Rate Before {i}"
        ama_columns_mapping[f"Số tiền thuế khác {i} trước khi sửa"] = f"Other Tax Amount Before {i}"
        ama_columns_mapping[f"Miễn thuế khác {i} trước khi sửa"] = f"Other Tax Exemption Before {i}"
        ama_columns_mapping[f"Trị giá TT khác {i} sau khi sửa"] = f"Other Tax Value After {i}"
        ama_columns_mapping[f"Lượng TT khác {i} sau khi sửa"] = f"Other Quantity After {i}"
        ama_columns_mapping[f"Mã DVT thuế khác {i} sau khi sửa"] = f"Other Unit After {i}"
        ama_columns_mapping[f"Mã áp dụng thuế khác {i} sau khi sửa"] = f"Other Tax Code After {i}"
        ama_columns_mapping[f"Thuế suất khác {i} sau khi sửa"] = f"Other Tax Rate After {i}"
        ama_columns_mapping[f"Số tiền thuế khác {i} sau khi sửa"] = f"Other Tax Amount After {i}"
        ama_columns_mapping[f"Miễn thuế khác {i} sau khi sửa"] = f"Other Tax Exemption After {i}"
        ama_columns_mapping[f"Số tiền điều chỉnh thuế khác {i}"] = f"Adjusted Other Tax Amount {i}"
        ama_columns_mapping[f"Mã nguyên tệ DC thuế khác {i}"] = f"Adjusted Other Tax Currency {i}"

    df.columns = rename_duplicate_columns(df.columns)
    df = df.rename(columns=ama_columns_mapping)
    df['Update Date'] = date_of_checking
    df.loc[:, 'Original CDS'] = df['Original CDS'].astype(str)
    df.loc[:, 'Amendament Number'] = df['Amendament Number'].astype(str)
    df.loc[:, 'Tax Code'] = df['Tax Code'].astype(str)
    df.loc[:, 'Phone'] = df['Phone'].astype(str)

    df['Amendament Date'] = pd.to_datetime(df['Amendament Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['Original CDS Date'] = pd.to_datetime(df['Original CDS Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['Approval Date'] = pd.to_datetime(df['Approval Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['Expiry Date'] = pd.to_datetime(df['Expiry Date'], errors='coerce').dt.strftime('%Y-%m-%d')
    df['Inspection Complete Date'] = pd.to_datetime(df['Inspection Complete Date'], errors='coerce').dt.strftime('%Y-%m-%d')

    unique_years = df['Original CDS Date'].str[:4].unique()  # Lấy 4 ký tự đầu tiên của 'CDS Date' để làm năm
    ex_im = df['Original CDS'].astype(str).str[:1].unique()
    for each in ex_im:
        # print(each)
        dept = "IMPORT" if each == "1" else "EXPORT"
        for year in unique_years:
            filename = f"AMA {dept} REPORT {year}.xlsx"
            # year_list.append(year)
            year_df = df[df['Original CDS Date'].str.startswith(year)]
            ex_im_df = year_df[year_df['Original CDS'].str.startswith(each)]
            # print(ex_im_df['Original CDS'])
            # Kiểm tra xem file đã tồn tại chưa
            if os.path.exists(filename):
                # Nếu file đã tồn tại, đọc file cũ vào DataFrame
                existing_df = pd.read_excel(filename)
                existing_df_filled = existing_df.fillna("")
                ex_im_df_filled = ex_im_df.fillna("")
                # print(existing_df_filled.columns)
                # print(ex_im_df_filled.columns)
                # Reset index to avoid reindexing issues
                existing_df_filled = existing_df_filled.reset_index(drop=True)
                ex_im_df_filled = ex_im_df_filled.reset_index(drop=True)

                combined_df = pd.concat([existing_df_filled, ex_im_df_filled], ignore_index=True)
                # combined_df = combined_df.drop_duplicates(keep='last')  # 'last' giữ dòng mới nhất (dòng vừa thêm vào)


            else:
            # Nếu file chưa tồn tại, chỉ sử dụng dữ liệu mới
                combined_df = ex_im_df

            combined_df.loc[:, 'Original CDS'] = combined_df['Original CDS'].astype(str)
            combined_df.loc[:, 'Amendament Number'] = combined_df['Amendament Number'].astype(str)
            combined_df.loc[:, 'Tax Code'] = combined_df['Tax Code'].astype(str)
            combined_df.loc[:, 'Phone'] = combined_df['Phone'].astype(str)
            combined_df = combined_df.drop_duplicates(
                subset=["Amendament Number", "Original CDS", "Material Code Before", "Items Lines"],
                keep="last"
            )
            combined_df.to_excel(filename, index=False, sheet_name="AMA")

            print(f"File for year {year} saved as {filename}")
            if "EXPORT" in filename:
                destination_directory = r"\\10.20.254.8\Data\Dept\Admin-Shipping\0-部門公共檔案 SH Public\CDS\AMA\Export"
            elif "IMPORT" in filename:
                destination_directory = r"\\10.20.254.8\Data\Dept\Admin-Shipping\0-部門公共檔案 SH Public\CDS\AMA\Import"
            else:
                destination_directory = r"\\10.20.254.8\Data\Dept\Admin-Shipping\0-部門公共檔案 SH Public\CDS"
            shutil.copy(filename, destination_directory)
            print("Copy the file to destination")
            error_list.append(f"Success AMA")
def running(date_of_checking):
    try:
        # print(today)
        error_list.append(f"Running extract file: EXPORT {date_of_checking}\n"
                          f"Running extract file: IMPORT {date_of_checking}\n"
                          f"Running extract file: AMA REPORT {date_of_checking}"
                          )
        subject_keyword = f"[RPA] DOWNLOADING CDS REPORT {date_of_checking}"  # Thay thế bằng chủ đề bạn đang tìm
        save_folder = r'\\10.20.254.8\Data\Dept\Admin-Shipping\0-部門公共檔案 SH Public\CDS\RPA DATA'  # Thư mục bạn muốn lưu tệp đính kèm
        save_attachments(subject_keyword, save_folder)
        extractfile(name=f"EXPORT {date_of_checking}")
        extractfile(name=f"IMPORT {date_of_checking}")
        extractfile(name=f"AMA REPORT {date_of_checking}")
    except Exception as e:
        error_list.append(f"Something wrong: {e}")
        print(f"Something wrong: {e}")
        sendmail()
# for month in ["01","02","03"]:
#     for date in range(1,32):
#         if date < 10:
#             date = f"0{date}"
#         date_of_checking = f"2025-{month}-{date}"
#         print(date_of_checking)
#         # extractfile(name=f"EXPORT {date_of_checking}")
#         # extractfile(name=f"IMPORT {date_of_checking}")
#         extractfile(name=f"AMA REPORT {date_of_checking}")
# #
# year_list = ["2018","2019","2020","2021","2022","2023","2024","2025"]
# combine_file2()

# date_of_checking = "2025-01-17" - cds 2018-2025
# date_of_checking = "2025-03-05" # 2017-2025 - AMa
# extractfile(name=f"AMA REPORT {date_of_checking}")
#
# extractfile(name=f"EXPORT {date_of_checking}")
# extractfile(name=f"IMPORT {date_of_checking}")

# date_of_checking = today
# running()

for i in reversed(range(10)):  # chạy từ 9 → 0
    date_obj = datetime.today() - timedelta(days=i)
    date_of_checking = date_obj.strftime('%Y-%m-%d')
    print(f"🔍 Đang chạy kiểm tra ngày: {date_of_checking}")
    checking_date_list.append(date_of_checking)
    running(date_of_checking)  # Nếu running() nhận ngày làm đối số
combine_file2()
# sendmail()
