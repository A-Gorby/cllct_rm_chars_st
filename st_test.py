import streamlit as st
import pandas as pd
from utils_main import read_okpd_dict_fr_link, split_merged_cells_st, extract_spgz_df_lst_st
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
if len (fn_lst) == 0:
    st.write(f"В загруженных файлах не найдены .xlsx файлы")
    # st.write(f"Работа программы завершена")
    # st.write(f"Обновите страницу")
    # sys.exit(2)
else:
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

            spgz_code_name, spgz_characteristics_content_loc_df = extract_spgz_df_lst(
              fn=fn_proc_save,
              sh_n_spgz=sh_n_source,
              groupby_col='№п/п',
              unique_test_cols=['Наименование СПГЗ', 'Единица измерения', 'ОКПД 2', 'Позиция КТРУ'],
              significant_cols = [
                  'Наименование характеристики', 'Единица измерения характеристики', 'Значение характеристики', 'Тип характеристики', 'Тип выбора значений характеристики заказчиком'],
            )
            if debug: print(spgz_code_name)
            #     kpgz_head, chars_of_chars_df = create_kpgz_data(spgz_characteristics_content_loc_df, debug = False)

        #     fn_save = fn_source.split('.xlsx')[0] + '_upd.xlsx'
        #     write_head_kpgz_sheet(
        #         data_source_dir,
        #         data_processed_dir,
        #         fn_source,
        #         fn_save,
        #         spgz_code_name,
        #         kpgz_head,
        #         chars_of_chars_df,
        #         okpd2_df,
        #         debug=False
        #     )