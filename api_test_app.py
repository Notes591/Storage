# -*- coding: utf-8 -*-
# تطبيق تجربة الـ API — نفس stock_requests.py بالظبط، بس بيقرا الطلبات والمخزون من تابات الـ API.
# في Streamlit Cloud: New app من نفس الريبو، Main file = api_test_app.py (محتاج نفس secrets الأساسي).
# ⚠️ الملف ده لازم يتحفظ باسم api_test_app.py — مش stock_requests.py.
import os
import runpy

_here = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(__file__).lower() == "stock_requests.py":
    import streamlit as st
    st.error("الملف ده (التجربة) اتسمّى stock_requests.py بالغلط وكتب فوق التطبيق الأساسي. "
             "ارفع stock_requests.py الكامل تاني، وسمّي ملف التجربة api_test_app.py.")
    st.stop()

os.environ["ORDERS_TAB"] = "DailyOrders_API"
os.environ["INVENTORY_TAB"] = "Inventory_API"
runpy.run_path(os.path.join(_here, "stock_requests.py"), run_name="__main__")
