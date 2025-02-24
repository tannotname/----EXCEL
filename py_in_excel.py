import openpyxl
import tkinter as tk
from tkinter import messagebox

def find_and_update(file_path, search_value, add_hours, record_id, history_text, entry_id, entry_hours, entry_record):
    wb = openpyxl.load_workbook(file_path)
    
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=False):
            student_id = str(row[2].value).strip()  # C欄學號（轉為字串）
            student_name = str(row[3].value).strip()  # D欄姓名
            
            if search_value.strip() in [student_id, student_name]:
                # 更新 F 欄（時數累加）
                current_hours = row[5].value or 0
                row[5].value = current_hours + add_hours
                
                # 尋找 J 欄以後的空白欄位，寫入登記編號
                col_index = 9  # J 欄的索引是 9（從 0 計算）
                while col_index < sheet.max_column and row[col_index].value:
                    col_index += 1
                
                # 如果欄位不夠，自動增加
                if col_index >= sheet.max_column:
                    sheet.insert_cols(col_index + 1)
                row[col_index].value = record_id
                
                wb.save(file_path)
                wb.close()
                messagebox.showinfo("成功", f"更新成功：{search_value}\n時數 +{add_hours}\n登記編號 {record_id}")
                
                # 記錄歷史紀錄
                history_text.insert(tk.END, f"{search_value} | +{add_hours} 小時 | 編號: {record_id}\n")
                history_text.yview(tk.END)
                
                # 清空輸入欄位
                entry_id.delete(0, tk.END)
                entry_hours.delete(0, tk.END)
                entry_record.delete(0, tk.END)
                return
    
    wb.close()
    messagebox.showwarning("錯誤", "未找到符合條件的學生。")

def submit_data(entry_id, entry_hours, entry_record, history_text, file_path):
    search_value = entry_id.get().strip()
    add_hours = entry_hours.get().strip()
    record_id = entry_record.get().strip()
    
    if not search_value or not add_hours or not record_id:
        messagebox.showwarning("輸入錯誤", "請填寫所有欄位！")
        return
    
    try:
        add_hours = int(add_hours)
        find_and_update(file_path, search_value, add_hours, record_id, history_text, entry_id, entry_hours, entry_record)
    except ValueError:
        messagebox.showwarning("格式錯誤", "時數請輸入數字！")

def main():
    file_path = "113-1-1生活記點名單.xlsx"
    
    root = tk.Tk()
    root.title("學生時數登記系統")
    
    tk.Label(root, text="學號或姓名:").grid(row=0, column=0)
    entry_id = tk.Entry(root)
    entry_id.grid(row=0, column=1)
    
    tk.Label(root, text="增加的時數:").grid(row=1, column=0)
    entry_hours = tk.Entry(root)
    entry_hours.grid(row=1, column=1)
    
    tk.Label(root, text="登記編號:").grid(row=2, column=0)
    entry_record = tk.Entry(root)
    entry_record.grid(row=2, column=1)
    
    history_text = tk.Text(root, height=10, width=50)
    history_text.grid(row=4, column=0, columnspan=2)
    tk.Label(root, text="歷史紀錄:").grid(row=3, column=0, columnspan=2)
    
    submit_button = tk.Button(root, text="提交", command=lambda: submit_data(entry_id, entry_hours, entry_record, history_text, file_path))
    submit_button.grid(row=5, column=0, columnspan=2)
    
    root.mainloop()
    
if __name__ == "__main__":
    main()
