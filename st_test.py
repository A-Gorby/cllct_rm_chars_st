import streamlit as st
import pandas as pd
from utils_main import read_okpd_dict_fr_link, split_merged_cells_st, extract_spgz_df_lst_st, write_head_kpgz_sheet_st
from utils_main import create_kpgz_data
import sys

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
sh_n_source = 'СПГЗ'
debug=False
debug=True

okpd2_df = read_okpd_dict_fr_link()
st.dataframe(okpd2_df.head(2)) #, use_container_width=True)

uploaded_files = st.file_uploader(
    "Загрузите xlsx- файлы для обработки", accept_multiple_files=True
)
fn_lst = [fn.name for fn in  uploaded_files if fn.name.endswith('.xlsx')]
fn_save_lst = []
if len (fn_lst) == 0:
    st.write(f"В загруженных файлах не найдены .xlsx файлы")
    # st.write(f"Работа программы завершена")
    # st.write(f"Обновите страницу")
    # sys.exit(2)
else:
    fn_save_lst = []
    for uploaded_file in uploaded_files:
        # bytes_data = uploaded_file.read()
        # st.write("filename:", uploaded_file.name)
        # st.write(bytes_data)
        if uploaded_file.name.endswith('.xlsx'):
            # try:
            #     df = pd.read_excel(uploaded_file)
            #     st.dataframe(df, use_container_width=True)
            # except:      
            #     df=pd.DataFrame()

            # fn_proc_save = split_merged_cells(fn_path, sh_n_spgz=sh_n_source, save_dir=data_tmp_dir, debug=False)
            fn_proc_save = split_merged_cells_st(uploaded_file, sh_n_spgz=sh_n_source, save_suffix='_spliited', debug=False)

            ##     df_rm_source = read_data(data_tmp_dir, fn_source, sh_n_source, )

            spgz_code_name, spgz_characteristics_content_loc_df = extract_spgz_df_lst_st(
              fn=fn_proc_save,
              sh_n_spgz=sh_n_source,
              groupby_col='№п/п',
              unique_test_cols=['Наименование СПГЗ', 'Единица измерения', 'ОКПД 2', 'Позиция КТРУ'],
              significant_cols = [
                  'Наименование характеристики', 'Единица измерения характеристики', 'Значение характеристики', 'Тип характеристики', 'Тип выбора значений характеристики заказчиком'],
            )
            if debug: 
                st.write(spgz_code_name)
                st.dataframe(spgz_characteristics_content_loc_df.head(2))
            
            
            kpgz_head, chars_of_chars_df = create_kpgz_data(
                spgz_characteristics_content_loc_df, debug = False)
            if debug: 
                st.write(kpgz_head)
                st.dataframe(chars_of_chars_df.head(2))

            fn_save = uploaded_file.name.split('.xlsx')[0] + '_upd.xlsx'
            write_head_kpgz_sheet_st(
                    uploaded_file,
                    fn_save,
                    spgz_code_name,
                    kpgz_head,
                    chars_of_chars_df,
                    okpd2_df,
                    debug=False
                )
            fn_save_lst.append (fn_save)

            # ---
# Binary files
import zipfile

fn_zip = "form_spgz.zip"
with zipfile.ZipFile(fn_zip, "a") as zf:
    fn_save_lst - list(set(fn_save_lst))
    for fn_save in fn_save_lst:
        zf.write(fn_save)
        # break
    st.write(zf.namelist())
# binary_contents = b'whatever'

# Different ways to use the API

# if st.download_button('Download file', fn_zip):  # Defaults to 'application/octet-stream'

with open(fn_zip, 'rb') as f:
    if st.download_button('Download Zip', f, mime='application/octet-stream', file_name=fn_zip):  # Defaults to 'application/octet-stream'


        # You can also grab the return value of the button,
        # just like with any other button.
        st.write('Thanks for downloading!')

#     # if st.download_button(...):
#     #     st.write('Thanks for downloading!')