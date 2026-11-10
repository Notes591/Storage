# -*- coding: utf-8 -*-
# تطبيق تجربة الـ API — نفس stock_requests.py بالظبط، بس بيقرا الطلبات والمخزون من تابات الـ API.
# في Streamlit Cloud: اعمل تطبيق جديد من نفس الريبو وخلّي Main file = stock_requests_api.py
# (محتاج نفس secrets بتاعة التطبيق الأساسي، ومفيش أي سطر زيادة).
import os
import runpy

os.environ["ORDERS_TAB"] = "DailyOrders_API"
os.environ["INVENTORY_TAB"] = "Inventory_API"

runpy.run_path(os.path.join(os.path.dirname(os.path.abspath(__file__)), "stock_requests.py"), run_name="__main__")
