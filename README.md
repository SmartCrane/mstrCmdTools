# mstrCmdTools
Command Manager Tools Edit by Python 3.
author:Anning
@(06-00-01 综合技术)[Microstrategy，command manager,Python]
####使用示例代码：
    
    from Tools.MstrTools.mstrCmdTools import MicrostrategyCmdTools 
    UserName_mstr_tools = MicrostrategyCmdTools('cmdmgr', 'UserName', '\'\'', 'ProjectName')
    user_address_script_path = '... .../dealWithUserAddress.scp'
    sub_modify_script_path = '... .../modifysub.scp'
    user_status_script_path = '... .../dealWithUserStatus.scp'
    update_repot_cache_script_path = '... .../updateReportCache.scp'
    excel_delivery_format = 'EXCEL'
    
    使用示例：
    def generate_report():
    '''
    生成报表
    :return:
    '''

    try:
        
        '''启用UserName用户'''
        set_user_status_script = admin_mstr_tools.set_user_status_script('UserName', 'ENABLED')
        print(set_user_status_script)
        admin_mstr_tools.exec_cmg_script(set_user_status_script, user_status_script_path)
        # 连接数据库
        db_ods = OraDB('xxx', 'xxx', 'xxx', 'xxx')
        '''
        遍历
        '''
        name_sql = '''
            select distinct
               name,
               report_name,
               to_number(substr(name, 1, instr(name, '、', 1, 1) - 1)) rn,
               dept_name
          from table_name
         where stat_date = ''' + month_num + '''
           and name is not null
        order by to_number(substr(name, 1, instr(name, '、', 1, 1) - 1))
        '''
        print(name_sql)

        name_rows = db_ods.exec_sql(name_sql)

        for row in name_rows:
            try:
                '''
                生成报表
                '''
                address_name = row[0]
                physical_address = row[0]
                # 1、处理User Address
                user = 'UserName'
                create_user_address_script = admin_mstr_tools.create_user_address_script(address_name, user, physical_address, device)
                admin_mstr_tools.exec_cmg_script(create_user_address_script, user_address_script_path)

                dept_name = str(row[3])
                # generate_file_name = row[1]
                generate_file_name = year_num + '年' + month_rn + '月' + two_dept_name + '报表'

                answer = dept_name
                answer_list = answer.split(',')
                answer_exp = ",".join('^\"' + answer_element + '^\"' for answer_element in answer_list)
                prompt_answer = '[\框架对象\实体\\组织结构\报表\部门]@ID in (' + answer_exp + ')'

                prompts = [{'promptName': '部门条件限定提示', 'answer': prompt_answer},
                           {'promptName': '月份', 'answer': month_num}]

                file_subscription_modify_script = UserName_mstr_tools.excel_subscription_script(
                    subscription_name,
                    user,
                    schedule_name,
                    address_name,
                    REPORT_GUID,
                    generate_file_name,
                    prompts,
                    excel_delivery_format
                )
                print(file_subscription_modify_script)
                UserName_mstr_tools.exec_cmg_script(file_subscription_modify_script, sub_modify_script_path)
            except Exception as err:
                print(err)
                continue

    except Exception as err:
        print(err)
    finally:
        # 断开数据库
        db_ods.close()
        '''禁用UserName用户'''
        set_user_status_script = admin_mstr_tools.set_user_status_script('UserName', 'DISABLED')
        admin_mstr_tools.exec_cmg_script(set_user_status_script, user_status_script_path)
