# -*- coding: utf-8 -*-

import sys
import xlrd
import xlsxwriter
from netaddr import IPNetwork, IPAddress

# writing one row in excel table
def row_write(row, hostname, group, int, proto, p_src, ip_dst, p_dst, ip_src, ip_dst_azs=None):
    global worksheet
    worksheet.write_string(row, 0, hostname)
    worksheet.write_string(row, 1, group)
    worksheet.write_string(row, 2, int)
    worksheet.write_comment(row, 2, ip_src)
    worksheet.write_string(row, 3, proto)
    worksheet.write_string(row, 5, ip_dst)
    if ip_dst_azs is not None:
        worksheet.write_string(row, 5, ip_dst, cell_bold_cyan)
        worksheet.write_comment(row, 5, ip_dst_azs)

    if dynamic in p_src:
        p_src[p_src.index(dynamic)] = 'DYN'
        worksheet.write_comment(row, 4, 'DYN здесь подразумевается порты выше '+str(dynamic))

    if dynamic in p_dst:
        p_dst[p_dst.index(dynamic)] = 'DYN'
        worksheet.write_comment(row, 6, 'DYN здесь подразумевается порты выше '+str(dynamic))

    if proto != 'icmp':
        worksheet.write_string(row, 4, ', '.join(map(lambda x: str(x), p_src)))
        worksheet.write_string(row, 6, ', '.join(map(lambda x: str(x), p_dst)))
    else:
        worksheet.write_string(row, 4, ' --- ')
        worksheet.write_string(row, 6, ' --- ')

# split line to list
def split_line(line_str):
    line_list = line_str.split()
    if line_list[0] == 'icmp':
         return line_list[0], line_list[1].split('=')[1], line_list[2].split('=')[1], 0, 0
    else:
         return line_list[0], line_list[1].split('=')[1], line_list[2].split('=')[1], dynamic if int(line_list[3].split('=')[1]) > dynamic else int(line_list[3].split('=')[1]), int(line_list[4].split('=')[1])

# script location
current_folder = sys.path[0]
# file with conntrack dump
link = current_folder + '/conntrack.txt'
# counter of unique sessions
counter = 0
# dynamic ports starts from...
dynamic = 24000

# resulted list of unique connections
clear_list = []

print('> Summarizing ports')

# open conntrack file
with open(link, 'r') as file:
    for line in file.readlines():
        # get protocol, IPs and ports of source and destination
        proto, ip_src, ip_dst, p_src, p_dst = split_line(line)
        
        # try to find in resulted list and summarizing ports
        flag_find = False
        for clean in clear_list:
            if (clean['ip_src'] == ip_src) and (clean['ip_dst'] == ip_dst) and (clean['proto'] == proto):

                # if source port < then dynamic port adding to list
                if (p_src not in clean['p_src']) and (p_src < dynamic):
                     clean['p_src'].append(p_src)

                # if destination port < then dynamic port adding to list
                if (p_dst not in clean['p_dst']) and (p_dst < dynamic):
                     clean['p_dst'].append(p_dst)

                flag_find = True
                break

        # add to resulted list new connection
        if flag_find is False:
            if counter % 100 == 0:
                print('!', end='', flush=True)
            counter += 1
            clear_list.append({'proto': proto, 'ip_src': ip_src, 'ip_dst': ip_dst, 'p_src': [p_src], 'p_dst': [dynamic] if p_dst > dynamic else [p_dst], 'int': 'None', 'host': 'None', 'group': 'None'})


print('')

# now we need to find interfaces of source and destination ip and hostnames
# file with hosts with subnets on interfaces 
xlsx_link = current_folder + '/subnets.xlsx'
xlsx = xlrd.open_workbook(xlsx_link)
sheet = xlsx.sheet_by_index(0)
# list of connections wich we finded
succ_list = []
# list of connections wich we didn't find
fail_list = []

print('> Looking for interfaces in subnets.xlsx')

# for each connection
for line in clear_list:
    flag_finded = False
    # we try to find in subnets file
    for row_num in range(1, sheet.nrows):
        row = sheet.row_values(row_num)
        if row[2].strip() != '':
            # we find source IP address in subnet on host!
            if IPAddress(line['ip_src']) in IPNetwork(row[2]):
                # now we looking for interface with with IP
                for net in row[3:]:
                    if (net.strip() != '') and (IPAddress(line['ip_src']) in IPNetwork(str(IPAddress(net.split()[0])) + '/' + str(IPAddress(net.split()[1]).netmask_bits()))):
                        line['int'] = sheet.cell_value(0, row.index(net))
                        line['host'] = sheet.cell_value(row_num, 0)
                        line['group'] = sheet.cell_value(row_num, 1)
                        line['ip_src'] = str(IPNetwork(str(IPAddress(net.split()[0])) + '/' + str(IPAddress(net.split()[1]).netmask_bits())))
                        flag_finded = True
                        break
                break
            # we find destination IP address in subnet on host!
            elif IPAddress(line['ip_dst']) in IPNetwork(row[2]):
                # now we looking for interface with with IP
                for net in row[3:]:
                    if (net.strip() != '') and (IPAddress(line['ip_dst']) in IPNetwork(str(IPAddress(net.split()[0])) + '/' + str(IPAddress(net.split()[1]).netmask_bits()))):
                        line['ip_dst'], line['ip_src'] = line['ip_src'], str(IPNetwork(str(IPAddress(net.split()[0])) + '/' + str(IPAddress(net.split()[1]).netmask_bits())))
                        line['p_dst'], line['p_src'] = line['p_src'], line['p_dst']
                        line['int'] = sheet.cell_value(0, row.index(net))
                        line['host'] = sheet.cell_value(row_num, 0)
                        line['group'] = sheet.cell_value(row_num, 1)
                        flag_finded = True
                        break
                break
    if flag_finded is False:
        # if we dont find such IPs in subnets list, adding with session to fail_list,
        # we will write it to second Excel Worksheet  
        fail_list.append(line)

print('> Summarizing interfaces')

# now we neet to summarizing IPs if it in one subnet
for line_num in range(len(clear_list)-1):
    if clear_list[line_num]['int'] != 'None':
         temp_list = clear_list[line_num]
         for x_line_num in range(line_num+1, len(clear_list)):
             if (temp_list['host'] == clear_list[x_line_num]['host']) and (temp_list['int'] == clear_list[x_line_num]['int']) and (temp_list['ip_dst'] == clear_list[x_line_num]['ip_dst']) and (temp_list['proto'] == clear_list[x_line_num]['proto']):
                 temp_list['p_src'] = list(set(temp_list['p_src']+clear_list[x_line_num]['p_src']))
                 temp_list['p_dst'] = list(set(temp_list['p_dst']+clear_list[x_line_num]['p_dst']))
                 clear_list[x_line_num]['int'] = 'None'
         succ_list.append(temp_list)

print('> Writing!')

# counter for rows
row_counter = 0
# open new xlsx
with xlsxwriter.Workbook(current_folder + '/conntrack.xlsx') as workbook:
    worksheet = workbook.add_worksheet('Sessions')

    cell_bold = workbook.add_format({'bold': True})
    cell_bold_cyan = workbook.add_format({'bold': True, 'bg_color': '#deff9e'})
    cell_bold_dark = workbook.add_format({'bold': True, 'bg_color': '#a4c661'})
    cell_bold_yellow = workbook.add_format({'bold': True, 'bg_color': '#d3e57b'})
    worksheet.set_column('A:A', None, cell_bold_cyan)
    worksheet.set_column('B:B', None, cell_bold_dark)
    worksheet.set_row(0, None, cell_bold_yellow)
    worksheet.set_column(0, 0, 25)
    worksheet.set_column(1, 1, 7)
    worksheet.set_column(2, 2, 25)
    worksheet.set_column(3, 3, 10)
    worksheet.set_column(4, 4, 80)
    worksheet.set_column(5, 5, 20)
    worksheet.set_column(6, 6, 80)

    worksheet.write_string(row_counter, 0, 'Hostname')
    worksheet.write_string(row_counter, 1, 'Group')
    worksheet.write_string(row_counter, 2, 'Interface')
    worksheet.write_string(row_counter, 3, 'Protocol')
    worksheet.write_string(row_counter, 4, 'Source ports')
    worksheet.write_string(row_counter, 5, 'Destination')
    worksheet.write_string(row_counter, 6, 'Dest. ports')
    
    for line in succ_list:
        flag_find = False
        row_counter += 1
        for row_num in range(1, sheet.nrows):
            row = sheet.row_values(row_num)
            if (row[2].strip() != '') and (IPAddress(line['ip_dst']) in IPNetwork(row[2])):
                row_write(row_counter, line['host'], line['group'], line['int'], line['proto'], sorted(line['p_src']), sheet.cell_value(row_num, 0), sorted(line['p_dst']), line['ip_src'], line['ip_dst'])
                flag_find = True
                continue
        if flag_find is False:
            row_write(row_counter, line['host'], line['group'], line['int'], line['proto'], sorted(line['p_src']), line['ip_dst'], sorted(line['p_dst']), line['ip_src'])

    # creating second worksheet with non-finded IPs in subnets table
    row_counter = 0
    worksheet = workbook.add_worksheet('Not finded') 

    worksheet.set_column(0, 0, 10)
    worksheet.set_column(1, 1, 14)
    worksheet.set_column(2, 2, 80)
    worksheet.set_column(3, 3, 14)
    worksheet.set_column(4, 4, 80)

    worksheet.write_string(row_counter, 0, 'Protocol')
    worksheet.write_string(row_counter, 1, 'IP source')
    worksheet.write_string(row_counter, 2, 'Port source')
    worksheet.write_string(row_counter, 3, 'IP dest')
    worksheet.write_string(row_counter, 4, 'Port dest')



    for line in fail_list:
        row_counter += 1
        worksheet.write_string(row_counter, 0, line['proto'])
        worksheet.write_string(row_counter, 1, line['ip_src'])
        worksheet.write_string(row_counter, 2, ', '.join(map(lambda x: str(x), line['p_src'])))
        worksheet.write_string(row_counter, 3, line['ip_dst'])
        worksheet.write_string(row_counter, 4, ', '.join(map(lambda x: str(x), line['p_dst'])))

print('Counter =', counter)
