import os
import sys
import shutil

try:
    import pefile
except ImportError:
    print("安裝 pefile 模組中...")
    os.system(f"{sys.executable} -m pip install pefile")
    import pefile

def strip_signature(filepath):
    if not os.path.exists(filepath):
        print(f"找不到檔案: {filepath}")
        return False

    try:
        pe = pefile.PE(filepath)
        if not hasattr(pe, 'OPTIONAL_HEADER') or not hasattr(pe.OPTIONAL_HEADER, 'DATA_DIRECTORY'):
            print(f"檔案格式不正確: {filepath}")
            return False

        sec_dir = pe.OPTIONAL_HEADER.DATA_DIRECTORY[pefile.DIRECTORY_ENTRY['IMAGE_DIRECTORY_ENTRY_SECURITY']]
        
        if sec_dir.VirtualAddress == 0 and sec_dir.Size == 0:
            print(f"檔案已無數位簽章或狀態已清除: {filepath}")
            return True

        # 將憑證表位址與大小清零
        sec_dir.VirtualAddress = 0
        sec_dir.Size = 0

        # 將修改後的內容寫入暫存檔
        temp_file = filepath + ".tmp"
        pe.write(filename=temp_file)
        pe.close()

        # 覆蓋原檔案
        shutil.move(temp_file, filepath)
        print(f"成功清除數位簽章: {filepath}")
        return True

    except PermissionError:
        print(f"權限不足無法修改: {filepath}。請使用系統管理員權限執行或確保檔案未被使用。")
        return False
    except Exception as e:
        print(f"處理檔案時發生錯誤 {filepath}: {e}")
        return False

if __name__ == "__main__":
    print("=== 開始清除 Windows 7 DLL 相容性問題的數位簽章 ===")
    
    # 取得當前 Python 環境的 DLLs 路徑
    base_dir = os.path.dirname(sys.executable)
    dlls_dir = os.path.join(base_dir, "DLLs")
    
    # 目標要清除的檔案
    target_files = [
        "pyexpat.pyd",
        "_tkinter.pyd",
        "tcl86t.dll",
        "tk86t.dll",
        "libexpat.dll"
    ]
    
    for filename in target_files:
        filepath = os.path.join(dlls_dir, filename)
        if os.path.exists(filepath):
            strip_signature(filepath)

    # 檢查根目錄可能存在的 DLL
    for filename in target_files:
        filepath = os.path.join(base_dir, filename)
        if os.path.exists(filepath):
            strip_signature(filepath)

    print("=== 簽章清除作業結束 ===")
