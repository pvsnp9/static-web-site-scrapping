import urllib2
from bs4 import BeautifulSoup
import xlsxwriter
import re

global_key = 'Name'
key_list = ['Name', 'From', 'To', 'Model', 'Core Size', 'Header', 'Assembly', 'Auto Assembly', 'Manual Assembly', 'Core Number', 'Gaskets', 'Notes', 'Engine Number', 'Chassis Number', 'OEM Number', 'Image Url']
base_url = 'web url ' #classified 


def spider_head(root_url):
    link_list = []
    main_request = urllib2.Request(root_url, headers={'User-Agent': "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/534.30 (KHTML, like Gecko) Ubuntu/11.04 Chromium/12.0.742.112 Chrome/12.0.742.112 Safari/534.30"})
    main_content = urllib2.urlopen(main_request)
    main_soup = BeautifulSoup(main_content.read(), 'lxml')
    links = main_soup.find_all('table', attrs={'class': 'table'})
    for items in links:
        link_bs4set = items.find_all('a')
        for link in link_bs4set:
            link_list.append(link['href'])
    link_list = remove_duplicates_from_list(link_list)
    body = spider_body(link_list)
    spider = spider_legs(body)
    return spider


def spider_body(parts):
    body_links = []
    for part in parts:
        if not part == 'products/volkswagen': #this is not good way of doing thing, optimise it later
            body_link = base_url+part
            body_links.append(body_link.strip())
    return body_links


def spider_legs(legs):
    links = []
    for leg in legs:
        leg_request = urllib2.Request(leg, headers={'User-Agent': "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/534.30 (KHTML, like Gecko) Ubuntu/11.04 Chromium/12.0.742.112 Chrome/12.0.742.112 Safari/534.30"})
        leg_content = urllib2.urlopen(leg_request)
        leg_soup = BeautifulSoup(leg_content.read(), 'lxml')
        leg_branches = leg_soup.find_all('table', attrs={'class': 'table'})
        if len(leg_branches) < 1:
            #perform variant pattern
            #print 'This is null value i ma receiving'
            if has_nails(leg):
                links.append(leg.replace(base_url+'products', 'products'))
            else:
                fleshes = nail_flesh(leg)
                for flesh_link in fleshes:
                    links.append(flesh_link)
        else:
            for td_items in leg_branches:
                td_links = td_items.find_all('td', attrs={'valign': 'top'})
                for items in td_links:
                    td_href = items.find('a')
                    if td_href:
                        links.append(td_href['href'])
    return links


def has_nails(nails):#checking varied pattern returns true if it directly has product link, false if has different list
    url_parts = nails.split('/')
    product_name = url_parts[len(url_parts)-1].title()
    #####print product_name
    nail_request = urllib2.Request(nails, headers={'User-Agent': "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/534.30 (KHTML, like Gecko) Ubuntu/11.04 Chromium/12.0.742.112 Chrome/12.0.742.112 Safari/534.30"})
    nail_content = urllib2.urlopen(nail_request)
    nail_soup = BeautifulSoup(nail_content.read(), 'lxml')
    nail_name = nail_soup.find('table', attrs={'border': '0', 'cellspacing': '1', 'cellpadding': '1', 'align': 'left'})
    if nail_name is None:
        return False
    else:
        return True


def nail_flesh(find_flesh):
    flesh_links = []
    flesh_request = urllib2.Request(find_flesh, headers={'User-Agent': "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/534.30 (KHTML, like Gecko) Ubuntu/11.04 Chromium/12.0.742.112 Chrome/12.0.742.112 Safari/534.30"})
    flesh = urllib2.urlopen(flesh_request)
    flesh_soup = BeautifulSoup(flesh.read(), 'lxml')
    flesh_part = flesh_soup.find('table', attrs={'border': '0', 'cellspacing': '1', 'cellpadding': '1', 'width': '100%'})
    td_link = flesh_part.find_all('td', attrs={'valign': 'top'})
    for link in td_link:
        link_anchor = link.find_all('a')
        for item in link_anchor:
            flesh_links.append(item['href'])
    return flesh_links# return nested list


def remove_duplicates_from_list(my_list):
    my_list.sort()
    i = len(my_list) - 1
    while i > 0:
        if my_list[i] == my_list[i - 1] or my_list[i].startswith('http://'):
            my_list.pop(i)
        i -= 1
    #just to remoce blank and unwanted links
    my_list.pop(0)
    return my_list


def grab_source(url):
    result_set = []
    req = urllib2.Request(url, headers={'User-Agent' : "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/534.30 (KHTML, like Gecko) Ubuntu/11.04 Chromium/12.0.742.112 Chrome/12.0.742.112 Safari/534.30"})
    try:
        con = urllib2.urlopen(req)
    except urllib2.HTTPError, err:
        if err.code == 404:
            return
        else:
            pass
    soup = BeautifulSoup(con.read(), "lxml")
    root_element = soup.find_all('table', attrs={'border': '0', 'cellspacing': '1', 'cellpadding': '1', 'align': 'left'})
    image_tags = soup.find_all('img') ##image_td
    if not len(root_element) + 2 == len(image_tags):
        counter = 1
    else:
        counter = 2
    for tale_items in root_element:
        table_left = tale_items.find_all("td")##passes the whole top table
        raw_data = left_wing(table_left)
        temp_data = convert_to_list(raw_data)
        try:
            image_src = right_wing_get_image(image_tags[counter])
        except IndexError():
            pass
        counter += 1
        temp_data.append(image_src)
        result_set.append(temp_data)
    return result_set


def left_wing(table): #recieves any left table
    temp_data_set = 'Name:'
    for data in table:
        if temp_data_set == 'Name:':
            temp_data_set = temp_data_set + data.text
        elif data.text.endswith(':') or ':' in data.text:
            temp_data_set = temp_data_set + '|' + data.text
        else:
            if 'Enquire' in data.text:
                pass
            else:
                temp_data_set = temp_data_set + ' ' + data.text
    return temp_data_set


def right_wing_get_image(elements):
    try:
        server_path = 'Image Url:a1automotivecooling.co.nz/'+elements['src']
    except IndexError:
        server_path = "No image"
    return server_path


def convert_to_list(data_to_list):
    table_list = data_to_list.split('|')
    return table_list


def convert_to_dictionary(result_set_data):
    final_result_set = []
    for elements in result_set_data:
        temp_data = {}
        for i in elements:
            items = i.split(':')
            #regex to seperate date
            if get_key(items) == 'Name':
                name = format_name(get_value(items))
                (a, b) = (get_key(items), name)
                date_dict = (a, b)
                temp_data_dict = dict([date_dict])
                temp_data.update(temp_data_dict)

                date_range = re.findall(r'\D(\d{4})\D', get_value(items))#getting years in list
                #changing list to key 'from' and 'To' and push to main dict
                if len(date_range) == 1: #sometimes we have onwards date
                    (a, b) = ('To', date_range[0] + ' Onwards' )
                    date_dict = (a, b)
                    temp_data_dict = dict([date_dict])
                    temp_data.update(temp_data_dict)

                elif len(date_range) >= 2: #range
                    (c, d) = ('From', date_range[0])
                    date_dict = (c, d)
                    temp_data_dict = dict([date_dict])
                    temp_data.update(temp_data_dict)

                    (c, d) = ('To', date_range[1])
                    date_dict = (c, d)
                    temp_data_dict = dict([date_dict])
                    temp_data.update(temp_data_dict)

                else: #trash
                    pass
            else:
                (k, v) = (get_key(items).strip(), get_value(items))
                dict_item = (k,v)
                temp = dict([dict_item])
                temp_data.update(temp)
        final_result_set.append(temp_data)
    return final_result_set


def write_to_excel(result_list, key_list):
    workbook = xlsxwriter.Workbook('cssdata.xlsx')
    worksheet = workbook.add_worksheet()
    # Add a bold format to use to highlight cells.

    # Adjust the column width.
    worksheet.set_column(1, 1, 15)
    write_excel_headers(key_list, workbook, worksheet)
    row = 1
    for data_list in result_list:
        col = 0
        dict_keys = data_list.keys()
        #find two items in one and replicate that
        #double_name = data_list['Name'].strip().split('\n')
        #print double_name

        if find_new_properties_and_update_list(key_list, dict_keys):
            write_excel_headers(key_list, workbook, worksheet)

        for items in key_list:
            if items in data_list:
                worksheet.write_string(row, col, data_list[items].strip())
            else:
                worksheet.write_string(row, col, '')
            col += 1
        row += 1
    workbook.close()


def find_new_properties_and_update_list(program_list, dict_key_list):
    found = False
    for item in dict_key_list:
        if item in program_list:
            found = False
        else:
            key_list.append(item)
            found = True
    return found


def write_excel_headers(sorted_keys, workbook, worksheet):
    bold = workbook.add_format({'bold': 1})

    # Adjust the column width.
    worksheet.set_column(1, 1, 15)
    col_address = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    # Write some data headers.
    my_header = 0
    for items in sorted_keys:
        if my_header <= 25:
            worksheet.write(col_address[my_header]+'1', items, bold)
            level_one = 0
        elif my_header <= 50:
            worksheet.write(col_address[0] + col_address[level_one]+'1', items, bold)
            level_one += 1
            level_two = 0
        elif my_header <= 75:
            worksheet.write(col_address[1] + col_address[level_two]+'1', items, bold)
            level_two += 1
            level_three = 0
        elif my_header <= 100:
            worksheet.write(col_address[2] + col_address[level_three]+'1', items, bold)
            level_three += 1
        my_header += 1
        #worksheet.write(col_address[my_header]+'1', items, bold)
        #worksheet.write(col_address[my_header]+'1', items, bold)
        #my_header += 1


def get_key(key):
    return key[0]


def get_value(value):
    var = ''
    for j in range(1, len(value)):
        var = var + value[j]
    return var


def process_double_items(process_data):
    printable_data = []
    for unit_dict in process_data:
        unit_dict_keys = unit_dict.keys()
        if 'Auto Assembly' in unit_dict_keys and 'Manual Assembly' in unit_dict_keys:
            #duplicate the dict obj
            duplicate_item = unit_dict.copy()
            unit_dict['Name'] = format_name(unit_dict['Name'], 1, True)
            unit_dict['Manual Assembly'] = ''
            printable_data.append(unit_dict)
            #add new records to the dict
            duplicate_item['Name'] = format_name(duplicate_item['Name'], 1)
            duplicate_item['Auto Assembly'] = ''
            printable_data.append(duplicate_item)
        else:
            unit_dict['Name'] = format_name(unit_dict['Name'], 2) # format name for individual name
            printable_data.append(unit_dict)
    return printable_data


def format_name(names, index=0, flag=False):
    remove_new_line = names.strip('\n')
    cleaned = remove_new_line.strip(' ')
    name = cleaned.split(' ')
    if index == 1: #this should do split for two names
        single_name = cleaned.split(' ')
        if flag:
            return single_name[0]
        else:
            try:
                return single_name[1]
            except IndexError:
                return "Not found"
    elif index == 2: #for single item to remove years
        temp_name = names.split(' ')
        return temp_name[0]
    else: #this call can only suppposed to be made by convert to dict
        temp_name = ''
        try:
            temp_name = name[0] + ' ' + name[1]
        except IndexError:
            temp_name = name[0] + ' ' + '1997'
        return temp_name


if __name__ == "__main__":
    master_url = 'http://b2bvupnpujwfedppmjoh.dp.oa/qspevdut/joefy_ofx' # this is the target and its modified with a key  
    spider_legs = spider_head(master_url)
    print_data = []
    #print '+++++++++++++++++++++++++'
    #print len(spider_legs)
    for spider_leg in spider_legs:
        #print '+++++++++++++++++++++++++'
        #print 'for url: ' + spider_leg
        #print '+++++++++++++++++++++++++'
        #print 'Url: ' + base_url + spider_leg
        #print '+++++++++++++++++++++++++'
        leg_data = grab_source(base_url+spider_leg)
        if leg_data:
            leg_data_dict = convert_to_dictionary(leg_data)
            final_set = process_double_items(leg_data_dict)
            for final_set_items in final_set:
                print_data.append(final_set_items)
                print final_set_items
                print '******************************************************'
    print print_data
    write_to_excel(print_data, key_list)

