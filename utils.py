#%% Loading Necessary Data
import os 
from os.path import join
import time
import win32com.client

#%%
USER_ID = os.getlogin()
T2_MEASURES_DIRECTORY = f"C:/Users/{USER_ID}/OneDrive - Microsoft/General/T2 Metrix Database/Measures"
T2_REPORT_DIRECTORY = f"C:/Users/{USER_ID}/OneDrive - Microsoft/General/T2 Metrix Database/Report"

def timefunction(func):
    def wrapper(*args, **kwargs):
        start_time = time.time()
        output = func(*args, **kwargs)
        print(f"{func.__name__} takes {time.time() - start_time} seconds to finish.")
        return output

    return wrapper

@timefunction
def refresh_power_query(path):
    xl = win32com.client.DispatchEx("Excel.Application")
    wb = xl.workbooks.open(path)
    xl.Visible = True
    wb.RefreshAll()
    xl.CalculateUntilAsyncQueriesDone()
    wb.save()
    xl.Quit()

# %% Loading necessary file

#Need to revise the code based on below method.
#https://stackoverflow.com/questions/40893870/refresh-excel-external-data-with-python
# %%


refresh_power_query(join(T2_MEASURES_DIRECTORY,"MINMAX", "MinMax_2023-09-01.xlsx"))

# %%
