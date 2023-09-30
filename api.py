import requests,os,json,time
import pandas as pd
import win32com.client
ct = time.localtime()
em = str(ct.tm_min)
eh = str(ct.tm_hour)
ed = str(ct.tm_mday)
emm = str(ct.tm_mon)
ey = str(ct.tm_year)
if len(emm) < 2:
    emm = "0"+emm
class Orders:
    def __init__(self):
        self.destruct = False
        self.header = {
            "X-Auth-Token": "*****",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        self.parameters = {
            "min_date_created": f"{ey}-{emm}-{ed}T{eh-1}:30:00.000-00:00",
            "max_date_created": f"{ey}-{emm}-{ed}T{eh}:{em}:00.000-00:00"
        }
        orders_url = "****"
        r = requests.get(orders_url, headers=self.header,params=self.parameters).json()
        self.orders_df = pd.DataFrame(r)
        self.order_id = self.orders_df.iloc[:, 0]
    def getId(self):
        try:
            self.order_id[0] = self.order[0] - 1
        except:
            print("There is an order ID discrepancy, was there a server shutdown recently?")
            outlook = win32com.client.Dispatch('Outlook.Application')
            olmailitem = 0x0
            newmail = outlook.CreateItem(olmailitem)
            newmail.Subject = "API Import Error"
            newmail.to = "****", "****"
            newmail.Body = (
                "THe BigCommerce import to Fishbowl has failed"
            )
            newmail.Send()
        self.product_dfs = []
        for i in self.order_id:
            self.products_url = "****"
            r = requests.get(self.products_url,headers=self.header).json()
            self.product_df = pd.DataFrame(r)
            self.product_dfs.append(self.product_df)
            self.all_products_df = pd.concat(self.product_dfs, ignore_index=True)
            order_id = self.all_products_df["order_id"]
            self.variant_id = self.all_products_df["variant_id"]
            # self.variant_df = pd.DataFrame({"order_num": [order_id]},{"variant_num" :[self.variant_id]})
        self.size = len(self.product_df)
        self.addColumn()
    def addColumn(self):
        for i in range(self.size):
            column_name = f"variant_id_{+i}"
            self.orders_df[column_name] = self.variant_id[i]
        self.create()
    def create(self):
        self.orders_df = self.orders_df.drop(self.orders_df.columns[[
            3,4,5,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,26,27,29,31,34,36,
            37,38,39,40,41,42,43,44,45,49,50,51,52,53,55,56,57,58,59,60,61,62,63,64,65,
            66,67,68,69,70,71
        ]],axis=1)
        self.orders_df.to_csv("test.csv",index=False)
        self.destruct=True
        self.delete()
    def delete(self):
        my_file = r"C:\Users\DamienDavis\AppData\Roaming\JetBrains\PyCharmCE2023.1\scratches\test.csv"
        if os.path.isfile(my_file):
            if self.destruct==True:
                time.sleep(500)
                os.remove(r"C:\Users\DamienDavis\AppData\Roaming\JetBrains\PyCharmCE2023.1\scratches\test.csv")
o = Orders()
if __name__ in "__main__":
    if em == 30:
        o.getId()
