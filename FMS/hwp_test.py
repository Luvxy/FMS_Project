import os
import sys
import datetime
import time
import win32com.client as win32
import ctypes
import logging
import pyautogui as pa

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def creat_print_file(replace_pairs):
    if is_admin():
        try:
            file = "C:/Users/user/Desktop/2024 빵 인수증/C인수증.hwp"
            logging.debug(f"File path: {file}")

            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            logging.debug("HWP Object created")

            hwp.RegisterModule("FilePathCheckDLL", "AutomationKey")
            logging.debug("Module registered")
            
            hwp.Open(file, "HWP", "forceopen:true")
            logging.debug(f"File opened: {file}")
            
            for find_str, replace_str in replace_pairs:
                hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                hwp.HParameterSet.HFindReplace.FindString = find_str
                hwp.HParameterSet.HFindReplace.ReplaceString = replace_str
                hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
                hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                logging.debug(f"Replaced {find_str} with {replace_str}")
                        
            # Save the modified document
            today = str(datetime.date.today())
            save_path = f"C:/Users/user/Desktop/2024 빵 인수증/{today} 빵 인수증.hwp"
            hwp.SaveAs(save_path)
            logging.debug(f"File saved: {save_path}")

            # 대기 시간 추가
            time.sleep(3)

            # Reopen the saved document for printing
            hwp.Open(save_path, "HWP", "forceopen:true")
            logging.debug(f"File reopened for printing: {save_path}")

        except Exception as e:
            logging.error(f"Error: {e}")
        finally:
            try:
                hwp.Quit()
                logging.debug("HWP Quit successfully")
            except Exception as e:
                logging.error(f"Error quitting HWP: {e}")
    else:
        logging.debug("Script not running as admin, attempting to relaunch as admin")
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
