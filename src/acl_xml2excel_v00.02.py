import PySimpleGUI as sg
import pandas as pd
import xmltodict
import os
import datetime

def con_process(filepath):
    try:
        with open(filepath, "r") as fp:
            xml_data = fp.read()
        dict_data = xmltodict.parse(xml_data)
        symid = dict_data['SymCLI_ML']['Symmetrix']['Symm_Info']['symid']
    except:
        return ("Error : Invalid xml file. Please check the file.")
    else:
        df_ig = pd.DataFrame(index=[], columns=['MvName', 'IgName', 'Hba', 'HbaWwpn', 'Alias'], dtype=str)
        df_pg = pd.DataFrame(index=[], columns=['MvName', 'PgName', 'FaDir', 'FaPort', 'Wwpn'], dtype=str)
        df_sg = pd.DataFrame(index=[], columns=['MvName', 'SgName', 'CascadedSgName'], dtype=str)
        df_dev = pd.DataFrame(index=[], columns=['MvName', 'IgName', 'Symdev', 'Capacity(mb)', 'Capacity(gb)', 'HostLun'])
        df_dev = df_dev.astype({'Capacity(mb)': int, 'Capacity(gb)': float, 'HostLun': str})
        df_CascadedIgList = pd.DataFrame(index=[], columns=['CascadedIgName', 'Hba', 'HbaWwpn', 'Alias', 'MvName'], dtype=str)

        t01 = []
        for i01 in range(len(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'])):
            #print("{}/{}".format(i01, len(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'])))
            flg01 = False
            if dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators'] != None:
                for k01 in dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators']:
                    if k01 == 'group_name':
                        flg01 = True
                        break
            if not flg01:
                df01 = pd.DataFrame.from_dict(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01])
                df01 = df01.T
                df01 = df01.reset_index(drop=True)
                df01 = df01.drop(columns=['Device', 'Initiator_List', 'Initiators', 'SG_Child_info', 'Totals', 'last_updated', 'port_info'])
                view_name = df01['view_name'].values[0]
                ig_name = df01['init_grpname'].values[0]
                if "*" not in ig_name:
                    t01.append("{}@{}".format(view_name, ig_name))
                pg_name = df01['port_grpname'].values[0]
                sg_name = df01['stor_grpname'].values[0]

                if dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators'] == None:
                    df02 = pd.DataFrame([['', '', '']], columns=['wwn', 'user_node_name', 'user_port_name'])
                else:
                    if type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators']['wwn']) == str or type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators']['wwn']) == dict:
                        df02 = pd.DataFrame(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators'], index=[0], dtype=str)
                    else:
                        df02 = pd.DataFrame(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators'], dtype=str)
                if type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['port_info']['Director_Identification']) == str or type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['port_info']['Director_Identification']) == dict:
                    df03 = pd.DataFrame([dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['port_info']['Director_Identification']], dtype=str)
                else:
                    df03 = pd.DataFrame(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['port_info']['Director_Identification'], dtype=str)

                if dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['SG_Child_info']['child_count'] == '0':
                    df04 = pd.DataFrame([['', '']], columns=['group_name', 'Status'])
                else:
                    if type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['SG_Child_info']['SG']) == str or type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['SG_Child_info']['SG']) == dict:
                        df04 = pd.DataFrame([dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['SG_Child_info']['SG']], dtype=str)
                    else:
                        df04 = pd.DataFrame(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['SG_Child_info']['SG'], dtype=str)
                t03 = []
                if type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Device']) == str or type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Device']) == dict:
                    t04 = [dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Device']]
                else:
                    t04 = dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Device']
                for i02 in range(len(t04)):
                    df05 = pd.DataFrame(t04[i02], dtype=str)
                    t02 = [df05['dev_name'].unique()[0], int(df05['capacity'].unique()[0]), float(df05['capacity_gb'].unique()[0]), df05['host_lun'].unique()[0]]
                    t03.append(t02)
                df06 = pd.DataFrame(t03, columns=['Symdev', 'Capacity(mb)', 'Capacity(gb)', 'HostLun'])

                df02['MvName'] = view_name
                df02['IgName'] = ig_name
                if df02['user_node_name'].values[0] == "" and df02['user_port_name'].values[0] == "":
                    df02['Alias'] = ""
                else:
                    df02['Alias'] = df02['user_node_name'].str.cat(df02['user_port_name'], sep='/')
                df02 = df02.rename(columns={'user_port_name': 'Hba', 'wwn': 'HbaWwpn'})
                df_ig = pd.concat([df_ig, df02[['MvName', 'IgName', 'Hba', 'HbaWwpn', 'Alias']]])

                if "*" in ig_name:
                    df07 = pd.concat([df01[['view_name', 'init_grpname']], df02[['Hba', 'HbaWwpn', 'Alias']]], axis=1)
                    df07 = df07.fillna({'init_grpname': ''})
                    df07['CascadedIgName'] = df07['init_grpname'].apply(lambda x: str(x).replace(" *", ""))
                    df07 = df07.rename(columns={'view_name': 'MvName'})
                    df_CascadedIgList = pd.concat([df_CascadedIgList, df07[['CascadedIgName', 'Hba', 'HbaWwpn', 'Alias', 'MvName']]])

                df03['MvName'] = view_name
                df03['PgName'] = pg_name
                df03 = df03.rename(columns={'dir': 'FaDir', 'port': 'FaPort', 'port_wwn': 'Wwpn'})
                df_pg = pd.concat([df_pg, df03[['MvName', 'PgName', 'FaDir', 'FaPort', 'Wwpn']]])

                df04['MvName'] = view_name
                df04['SgName'] = sg_name
                df04 = df04.rename(columns={'group_name': 'CascadedSgName'})
                df_sg = pd.concat([df_sg, df04[['MvName', 'SgName', 'CascadedSgName']]])

                df06['MvName'] = view_name
                df06['IgName'] = ig_name
                df_dev = pd.concat([df_dev, df06])

        df_ig = df_ig.drop_duplicates()
        df_ig = df_ig.reset_index(drop=True)
        df_pg = df_pg.drop_duplicates()
        df_pg = df_pg.reset_index(drop=True)
        df_sg = df_sg.drop_duplicates()
        df_sg = df_sg.reset_index(drop=True)
        df_dev = df_dev.drop_duplicates()
        df_dev = df_dev.reset_index(drop=True)
        df_CascadedIgList = df_CascadedIgList.reset_index(drop=True)
        df_CascadedIgList = df_CascadedIgList.fillna({'MvName': ''})

        ################################
        df_mv = pd.DataFrame(index=[], columns=['MvName', 'IgName', 'Hba', 'HbaWwpn', 'Alias', 'CascadedIgName', 'PgName', 'FaDir', 'FaPort', 'Wwpn', 'SgName', 'CascadedSgName', 'Symdev', 'Capacity(mb)', 'Capacity(gb)', 'HostLun'])
        df_mv = df_mv.astype({'Capacity(mb)': int, 'Capacity(gb)': float, 'HostLun': str})

        for i01 in range(len(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'])):
            #print("{}/{}".format(i01, len(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'])))
            flg01 = False
            flg02 = False
            if dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators'] != None:
                for k01 in dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators']:
                    if k01 == 'group_name':
                        flg01 = True
                        break
            for i02 in t01:
                if i02.split('@')[0] == dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['view_name']:
                    flg02 = True
                    break

            if not flg01 and flg02:
                df01 = pd.DataFrame.from_dict(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01])
                df01 = df01.T
                df01 = df01.reset_index(drop=True)
                df01 = df01.drop(columns=['Device', 'Initiator_List', 'Initiators', 'SG_Child_info', 'Totals', 'last_updated', 'port_info'])
                view_name = df01['view_name'].values[0]
                ig_name = df01['init_grpname'].values[0]
                pg_name = df01['port_grpname'].values[0]
                sg_name = df01['stor_grpname'].values[0]

                df02 = df_ig[['Hba', 'HbaWwpn', 'Alias']][(df_ig['MvName'] == view_name) & (df_ig['IgName'] == ig_name)]
                df02 = df02.reset_index(drop=True)
                df03 = df_pg[['PgName', 'FaDir', 'FaPort', 'Wwpn']][(df_pg['MvName'] == view_name) & (df_pg['PgName'] == pg_name)]
                df03 = df03.reset_index(drop=True)
                df03.iloc[1:, 0] = ""
                df04 = df_sg[['SgName', 'CascadedSgName']][(df_sg['MvName'] == view_name) & (df_sg['SgName'] == sg_name)]
                df04 = df04.reset_index(drop=True)
                df05 = df_dev[['Symdev', 'Capacity(mb)', 'Capacity(gb)', 'HostLun']][(df_dev['MvName'] == view_name) & (df_dev['IgName'] == ig_name)]
                df05 = df05.reset_index(drop=True)
                df06 = pd.concat([df01[['view_name', 'init_grpname']], df02, df03, df04, df05], axis=1)
                df06 = df06.rename(columns={'view_name': 'MvName', 'init_grpname': 'IgName'})
                df_mv = pd.concat([df_mv, df06])

            if flg01 and not flg02:
                df01 = pd.DataFrame.from_dict(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01])
                df01 = df01.T
                df01 = df01.reset_index(drop=True)
                df01 = df01.drop(columns=['Initiator_List', 'Initiators', 'SG_Child_info', 'last_updated', 'port_info'])
                view_name = df01['view_name'].values[0]
                if type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators']['group_name']) == str or type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators']['group_name']) == dict:
                    df02 = pd.DataFrame([dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators']])
                else:
                    df02 = pd.DataFrame(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['Initiators'], dtype=str)
                ig_name = df02.iloc[0, 0]
                if type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['port_info']['Director_Identification']) == str or type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['port_info']['Director_Identification']) == dict:
                    df03 = pd.DataFrame([dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['port_info']['Director_Identification']])
                else:
                    df03 = pd.DataFrame(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['port_info']['Director_Identification'], dtype=str)
                if dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['SG_Child_info']['child_count'] == '0':
                    df04 = pd.DataFrame([['', '']], columns=['group_name', 'Status'])
                else:
                    if type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['SG_Child_info']['SG']) == str or type(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['SG_Child_info']['SG']) == dict:
                        df04 = pd.DataFrame([dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['SG_Child_info']['SG']])
                    else:
                        df04 = pd.DataFrame(dict_data['SymCLI_ML']['Symmetrix']['Masking_View'][i01]['View_Info']['SG_Child_info']['SG'], dtype=str)
                df05 = pd.concat(
                    [df01[['view_name', 'init_grpname']],
                     df02,
                     df01['port_grpname'],
                     df03,
                     df01['stor_grpname'],
                     df04['group_name']],
                    axis=1)
                df06 = df_dev[['Symdev', 'Capacity(mb)', 'Capacity(gb)', 'HostLun']][(df_dev['MvName'] == view_name) & (df_dev['IgName'] == "{} *".format(ig_name))]
                df06 = df06.reset_index(drop=True)
                df07 = pd.concat([df05, df06], axis=1)
                df07 = df07.set_axis(['MvName', 'IgName', 'CascadedIgName', 'PgName', 'FaDir', 'FaPort', 'Wwpn', 'SgName', 'CascadedSgName', 'Symdev', 'Capacity(mb)', 'Capacity(gb)', 'HostLun'], axis='columns')
                df_mv = pd.concat([df_mv, df07])

        df_mv = df_mv.fillna({'MvName': '', 'IgName': '', 'Hba': '', 'HbaWwpn': '', 'Alias': '', 'CascadedIgName': '', 'PgName': '', 'FaDir': '', 'FaPort': '', 'Wwpn': '', 'SgName': '', 'CascadedSgName': '', 'Symdev': '', 'Capacity(mb)': 0, 'Capacity(gb)': 0, 'HostLun': ''})
        df_mv = df_mv.astype({'Capacity(mb)': int, 'Capacity(gb)': float, 'HostLun': str})
        df_mv = df_mv.reset_index(drop=True)

        exportfile = "{}/mv_sid{}_{}.xlsx".format(os.path.dirname(filepath), symid, datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))
        df_mv.to_excel(exportfile, sheet_name='MaskingView', index=False)
        with pd.ExcelWriter(exportfile, engine="openpyxl", mode="a") as writer:
            df_CascadedIgList.to_excel(writer, sheet_name='CascadedIgList', index=False)
        return ("The process was successfully completed.\nexport file path : {}".format(exportfile))

def main():
    layout = [
        [sg.Text("xml file"), sg.InputText(), sg.FileBrowse(key="file")],
        [sg.Submit('Create excel file')],
        [sg.Text("""
    Operating Procedure
    1. Set the environment variables for SolutionsEnabler.
    (Windows)
     > set SYMCLI_OUTPUT_MODE=XML
    (Unix)
     # export SYMCLI_OUTPUT_MODE=XML

    2. Save the result of the symaccess command to a file.
     # symaccess -sid <SYMID> list view -detail > <FileName>.xml

    3. Select the xml file from "Browse" and click "Create excel file".

    4. An Excel file is generated in the same directory as the xml file.
            """, relief=sg.RELIEF_RIDGE, border_width=5 )],
    ]
    window = sg.Window('MaskingView xml2excel', layout, finalize=True)

    while True:
        event, values = window.read()
        if event == 'Create excel file':
            r = con_process(values['file'])
            sg.popup(r)
        if event == sg.WIN_CLOSED:
            break
    window.close()

if __name__ == '__main__':
    main()
