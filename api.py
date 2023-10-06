import requests,os,json,time,sys
import pandas as pd
import win32com.client
ct = time.localtime()
class Orders:
    if ct.tm_min == int(30):
        def __init__(self):
            self.destruct = False
            self.header = {
                "X-Auth-Token": "*****",
                "Content-Type": "application/json",
                "Accept": "application/json"
            }
            with open("prev_order_id.csv", "r") as f:
                f1 = f.read()
                self.min_order = f1[3:]
            self.parameters = {
                "min_id": self.min_order
            }
            orders_url = "******"
            try:
                self.r = requests.get(orders_url, headers=self.header,params=self.parameters).json()
                self.orders_df = pd.DataFrame(self.r)
                self.order_id = self.orders_df.iloc[:, 0]
                self.last_id = self.order_id[-1:]
                self.str_last_id = self.last_id.to_string(index=False)
                self.setId()
            except:
                self.email()
    else:
        raise RuntimeError

    def email(self):
        outlook = win32com.client.Dispatch('Outlook.Application')
        olmailitem = 0x0
        newmail = outlook.CreateItem(olmailitem)
        newmail.Sender = "Import Error"
        newmail.Subject = "API Import Error"
        newmail.to = "******"
        newmail.CC = "******"
        newmail.Body = (
            "The BigCommerce API has not imported any orders.\n"
            "If no orders were placed in the last hour, please disregard.\n\n"
            "If a new order was placed and the import failed, please check BigCommerce and the script."
        )
        newmail.Send()
        sys.exit(0)

    def setId(self):
            self.last_id=self.last_id+1
            self.last_id.to_csv("prev_order_id.csv", index=False)
            self.getId()

    def getId(self):
        self.product_dfs = []
        for i in self.order_id:
            self.products_url = "******"
            r = requests.get(self.products_url,headers=self.header).json()
            self.product_df = pd.DataFrame(r)
            self.product_dfs.append(self.product_df)
            self.all_products_df = pd.concat(self.product_dfs, ignore_index=True)
            self.order_id = self.all_products_df["order_id"]
            self.variant_id = self.all_products_df["variant_id"]
        self.size = len(self.variant_id)
        self.addColumn()

    def addColumn(self):
        for i in range(self.size):
            column_name = f"variant_id_{+i}"
            self.orders_df[column_name] = self.variant_id
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
        my_file = r"C:\Users\DamienDavis\Documents\API\test.csv"
        if os.path.isfile(my_file):
            if self.destruct==True:
                time.sleep(500)
                os.remove(r"C:\Users\DamienDavis\Documents\API\test.csv")

ct = time.localtime()
o = Orders()
if __name__ in "__main__":
    o.email()
