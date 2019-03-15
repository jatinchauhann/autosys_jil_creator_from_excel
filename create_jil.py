#!/usr/bin/env python
# coding: utf-8

# # JIL FILE CREATOR
# By Jatin Chauhan

import xlrd

#Change your file name here
file = ("file.xlsx")
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(0)

output = ""

for index in range(sheet.nrows):
    
    if sheet.cell_value(index,0) == 'box':
        
        row_data = sheet.row_values(index)
        
        props_box = {
            'job_type' : row_data[0],
            'box_name' : row_data[1],
            'description' : row_data[2],
            'owner' : row_data[3],
            'max_run_alarm' : row_data[4],
            'alarm_if_fail' : row_data[5],
            'date_conditions' : row_data[6],
            'send_notification' : row_data[7],
            'notification_msg' : row_data[8],
            'notification_emailaddress' : row_data[9]
        }
        
        output = output + '/* ----------------- {box_name} ----------------- */ \n\n'.format(**props_box)
        
        output = output + 'insert_job: {box_name} job_type: {job_type} \n'.format(**props_box)
        output = output + 'description: {description} \n'.format(**props_box)
        output = output + 'owner: {owner} \n'.format(**props_box)
        output = output + 'max_run_alarm: {max_run_alarm} \n'.format(**props_box)
        output = output + 'alarm_if_fail: {alarm_if_fail} \n'.format(**props_box)
        output = output + 'date_conditions: {date_conditions} \n'.format(**props_box)
        output = output + 'send_notification: {send_notification} \n'.format(**props_box)
        output = output + 'notification_msg: {box_name} {notification_msg} \n'.format(**props_box)
        output = output + 'notification_emailaddress: {notification_emailaddress} \n\n'.format(**props_box)
        
    if sheet.cell_value(index,0) == 'cmd' and not sheet.cell_value(index,15):
        
        row_data = sheet.row_values(index)
        
        props_cmd = {
            'job_type' : row_data[0],
            'box_name' : row_data[1],
            'description' : row_data[2],
            'owner' : row_data[3],
            'machine' : row_data[4],
            'max_run_alarm' : row_data[5],
            'alarm_if_fail' : row_data[6],
            'send_notification' : row_data[7],
            'notification_msg' : row_data[8],
            'notification_emailaddress' : row_data[9],
            'std_out_file' : row_data[10],
            'std_err_file' : row_data[11],
            'command' : row_data[12],
            'module_name' : row_data[13],
            'sub_module_name' : row_data[14]
        }
        
        output = output + '/* ----------------- {box_name} ----------------- */ \n\n'.format(**props_cmd)
        
        output = output + 'insert_job: {box_name} job_type: {job_type} \n'.format(**props_cmd)
        output = output + 'description: {description} \n'.format(**props_cmd)
        output = output + 'owner: {owner} \n'.format(**props_cmd)
        output = output + 'machine: {machine} \n'.format(**props_cmd)
        output = output + 'box_name: {box_name} \n'.format(**props_box)
        output = output + 'max_run_alarm: {max_run_alarm} \n'.format(**props_cmd)
        output = output + 'alarm_if_fail: {alarm_if_fail} \n'.format(**props_cmd)
        output = output + 'send_notification: {send_notification} \n'.format(**props_cmd)
        output = output + 'notification_msg: {box_name} {notification_msg} \n'.format(**props_cmd)
        output = output + 'notification_emailaddress: {notification_emailaddress} \n'.format(**props_cmd)
        output = output + 'std_out_file : {std_out_file}  \n'.format(**props_cmd)
        output = output + 'std_err_file : {std_err_file} \n'.format(**props_cmd)
        output = output + 'command : {command} {module_name} {sub_module_name} \n\n'.format(**props_cmd)
        
        props_cmd ={}
    
    if sheet.cell_value(index,0) == 'cmd' and sheet.cell_value(index,15):
        
        row_data = sheet.row_values(index)
        
        props_cmd = {
            'job_type' : row_data[0],
            'box_name' : row_data[1],
            'description' : row_data[2],
            'owner' : row_data[3],
            'machine' : row_data[4],
            'max_run_alarm' : row_data[5],
            'alarm_if_fail' : row_data[6],
            'send_notification' : row_data[7],
            'notification_msg' : row_data[8],
            'notification_emailaddress' : row_data[9],
            'std_out_file' : row_data[10],
            'std_err_file' : row_data[11],
            'command' : row_data[12],
            'module_name' : row_data[13],
            'sub_module_name' : row_data[14],
            'type_condition' : row_data[15],
            'condition' : row_data[16]
        }
        
        output = output + '/* ----------------- {box_name} ----------------- */ \n\n'.format(**props_cmd)
        
        output = output + 'insert_job: {box_name} job_type: {job_type} \n'.format(**props_cmd)
        output = output + 'description: {description} \n'.format(**props_cmd)
        output = output + 'owner: {owner} \n'.format(**props_cmd)
        output = output + 'machine: {machine} \n'.format(**props_cmd)
        output = output + 'box_name: {box_name} \n'.format(**props_box)
        output = output + 'condition: {type_condition}({condition}) \n'.format(**props_cmd)
        output = output + 'max_run_alarm: {max_run_alarm} \n'.format(**props_cmd)
        output = output + 'alarm_if_fail: {alarm_if_fail} \n'.format(**props_cmd)
        output = output + 'send_notification: {send_notification} \n'.format(**props_cmd)
        output = output + 'notification_msg: {box_name} {notification_msg} \n'.format(**props_cmd)
        output = output + 'notification_emailaddress: {notification_emailaddress} \n'.format(**props_cmd)
        output = output + 'std_out_file : {std_out_file}  \n'.format(**props_cmd)
        output = output + 'std_err_file : {std_err_file} \n'.format(**props_cmd)
        output = output + 'command : {command} {module_name} {sub_module_name} \n\n'.format(**props_cmd)
        
        props_cmd ={}
        

output_file = open(sheet.cell_value(0,1)+'.txt','w+')
output_file.write(output)
output_file.close()
output = '' 
