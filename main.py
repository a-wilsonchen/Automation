#%% 
import win32com.client
import time

#%%

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
    wb.Close(True)
    xl.Quit()
    
    

# %%
refresh_power_query("C:/Users/a-wilsonchen/OneDrive - Microsoft/General/T2 Metrix Database/Measures/DSM SUPPLY(Production+OH+Intransit)/DSM-SUPPLY_2023-09-01.xlsx")
# %%
