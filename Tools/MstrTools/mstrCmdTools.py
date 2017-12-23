#!/usr/local/bin/python3
# encoding=utf-8

"""
 2017-04-27, anning
"""
import os


class MicrostrategyCmdTools:

    def __init__(self, project_source, user_name, password, project):
        """
        Microstrategy CommandManager Tools 执行参数初始化
        :param project_source:
        :param user_name:
        :param password:
        :param project:
        """
        self.project_source = project_source
        self.user_name = user_name
        self.password = password
        self.project = project
        self.TRIGGER_SUBSCRIPTION_PATH = "/hdlroot/mstrDeliveriesFile/trigger_subscription.scp"
        self.mstr_bin_path = "/var/opt/MicroStrategy/bin"

    def exec_cmg_script(self, cmg_script, cmg_scp_script_path):
        """
        执行cmg处理脚本(通用方法)
        :param cmg_script:
        :param cmg_scp_script_path:
        :return:
        """
        with open(cmg_scp_script_path, "w+") as f:
            f.write(cmg_script)
        os.system("#!/bin/sh \n"
                  "cd %s \n"
                  "sudo ./mstrcmdmgr -n %s -u %s -p %s -f %s \n"
                  % (self.mstr_bin_path, self.project_source, self.user_name, self.password, cmg_scp_script_path))

    def set_user_status_script(self, user_group, status='DISABLED'):
        '''
        设置UserGroup的启用状态
        :param user_group:MicroStrategy 用户组
        :param status:用户启用状态，默认为DISABLED（禁用），ENABLED（启用）
        :return:
        '''
        deal_with_user_status_script = ""
        if status == 'DISABLED':
            deal_with_user_status_script = "ALTER USERS IN USER GROUP \'%s\' DISABLED;" % (user_group,)
        elif status == 'ENABLED':
            deal_with_user_status_script = "ALTER USERS IN USER GROUP \'%s\' ENABLED;" % (user_group,)

        return deal_with_user_status_script

    def create_user_address_script(self, address_name, user, physical_address, device):
        """
        生成Address处理脚本
        :param address_name:
        :param user:
        :param physical_address:
        :param device:
        :return:
        """
        deal_with_address_script = "REMOVE ADDRESS \"%s\"  FROM USER \"%s\"; \n " \
                                   "ADD ADDRESS \"%s\" PHYSICALADDRESS \"%s\" DELIVERYTYPE FILE DEVICE \"%s\" TO USER \"%s\"; \n" \
                                   % (address_name, user, address_name, physical_address, device, user)
        return deal_with_address_script

    def create_element_prompt_script(self, AttributeGUID, AnswerLists):
        '''
        :param AttributeGUID:
        :param AnswerLists:
        :return:
        '''
        element_prompt_script = ",".join(AttributeGUID + ':' + answer for answer in AnswerLists)
        # print(element_prompt_script)
        return element_prompt_script

    def excel_subscription_script(self, subscription_name, owner, schedule, address_name, report_service_document_guid,
                                  file_name, prompts, delivery_format):
        """
        生成 “用户订阅个性化内容，并激活订阅” 处理脚本
        :param subscription_name: MicroStrategy订阅名称
        :param owner:订阅Owner
        :param schedule:调度名称
        :param address_name: 用户的address_name;例如Administrator用户下的Address
        :param report_service_document_guid: 文档的ID；例如：B45FE6E711E74FE518FC0080EF858067
        :param file_name:生成文件的名称；
        :param prompts: [{'promptName':'年提示','answer':'2017'},{... ...}] 传入提示的顺序需要与文档的提示顺序一致
        :param delivery_format: EXECEL
        :return:
        """
        prompt_scripts_list = [
            "PROMPT \"" + prompts[count_num]['promptName'] + "\" ANSWER \"" + prompts[count_num]['answer'] + "\""
            for count_num in range(0, len(prompts)) if count_num < len(prompts)]

        prompt_scripts = ",".join(prompt for prompt in prompt_scripts_list)

        file_subscription_modify_script = "﻿DELETE SUBSCRIPTION \"%s\"  FROM PROJECT \"%s\"; \n" \
                                          "CREATE FILESUBSCRIPTION \"%s\" SCHEDULE \"%s\" " \
                                          "FOR OWNER \"%s\"  ADDRESS \"%s\" CONTENT GUID %s IN PROJECT \"%s\" DELIVERYFORMAT %s FILENAME \"%s\" " \
                                          "SENDPREVIEWNOW FALSE SENDNOTIFICATION FALSE %s; \n" \
                                          "TRIGGER SUBSCRIPTION \"%s\" FOR PROJECT \"%s\";" % (
                                              subscription_name, self.project, subscription_name, schedule, owner,
                                              address_name,
                                              report_service_document_guid, self.project, delivery_format, file_name,
                                              prompt_scripts,
                                              subscription_name, self.project
                                          )
        # print(file_subscription_modify_script)
        return file_subscription_modify_script

    def trigger_mstr_subscription(self, subscriptions):
        """
        根据传入的subscriptions，逐个触发subscription
        :param subscriptions:['subscription1','subscription2','subscription3']
        :return:None
        """
        for subscription_name in subscriptions:
            trigger_subscription_script = "TRIGGER SUBSCRIPTION \"%s\" FOR PROJECT \"%s\";" % (subscription_name, self.project)
            self.exec_cmg_script(trigger_subscription_script, self.TRIGGER_SUBSCRIPTION_PATH)

    def update_report_cache(self, subscription_name, schedule, user, report_guid, prompts):
        """
        更新Cash Subscription 更新报表缓存
        :param subscription_name:
        :param schedule:
        :param user:
        :param report_guid:
        :param prompts:
        :return:
        """
        prompt_scripts_list = ["PROMPT \"" + prompts[count_num]['promptName'] + "\" ANSWER \"" + prompts[count_num]['answer'] + "\""
                               for count_num in range(0, len(prompts)) if count_num < len(prompts)]

        prompt_scripts = ",".join(prompt for prompt in prompt_scripts_list)

        cache_subscription_modify_script = "﻿DELETE SUBSCRIPTION \"%s\"  FROM PROJECT \"%s\"; \n" \
                                           "CREATE CACHEUPDATESUBSCRIPTION \"%s\" SCHEDULE \"%s\" USER \"%s\" CONTENT GUID %s IN PROJECT \"%s\" " \
                                           "DELIVERYFORMAT HTML SENDNOTIFICATION FALSE %s; \n" \
                                           "TRIGGER SUBSCRIPTION \"%s\" FOR PROJECT \"%s\";" % (
                                               subscription_name, self.project, subscription_name, schedule, user,
                                               report_guid, self.project, prompt_scripts, subscription_name,
                                               self.project
                                            )
        print(cache_subscription_modify_script)
        return cache_subscription_modify_script
