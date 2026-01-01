import streamlit as st
st.write("# 看到這行字代表成功了！")
st.error("這是 v28 測試按鈕")
if st.button("點我清除快取"):
    st.cache_data.clear()
    st.write("快取已清除")
