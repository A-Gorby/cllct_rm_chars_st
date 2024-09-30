import streamlit as st
import pandas as pd
from utils_main import read_okpd_dict_fr_link

# df = pd.DataFrame(
#     [
#         {"command": "st.selectbox", "rating": 4, "is_widget": True},
#         {"command": "st.balloons", "rating": 5, "is_widget": False},
#         {"command": "st.time_input", "rating": 3, "is_widget": True},
#     ]
# )

# st.dataframe(df, use_container_width=True)
# import streamlit as st
# @st.cache_data
okpd2_df = read_okpd_dict_fr_link()
st.dataframe(okpd2_df[:5], use_container_width=True)
# uploaded_files = st.file_uploader(
#     "Choose a CSV file", accept_multiple_files=True
# )
# for uploaded_file in uploaded_files:
#     # bytes_data = uploaded_file.read()
#     # st.write("filename:", uploaded_file.name)
#     # st.write(bytes_data)

#     try:
#         df = pd.read_excel(uploaded_file)
#         st.dataframe(df, use_container_width=True)
#     except:      
#         df=pd.DataFrame()