from gooey import Gooey, GooeyParser

import xml.etree.ElementTree as ET
import csv
import re
import numpy as np
import pandas as pd
import os
def parse_all_text_tags_root(xml_content):
    root = ET.fromstring(xml_content)
    return root

def parse_all_text_tags_text(xml_content):
    root = ET.fromstring(xml_content)

    text_tags = root.findall('.//Text')
    
    texts = [text.text.strip() for text in text_tags if text.text is not None]

    return texts

def replace_tags(text, replacement_dict):
    pattern = r'Bus\[\d+\]\.Obj'
    if re.search(pattern, text):
        for key in replacement_dict.keys():
            text = text.replace(key,replacement_dict[key],1)
    return text

def replace_bus(text, replacement_dict,lst):
    pattern = r'Bus\[\d+\]\.Obj'
    if re.search(pattern, text):
        for i in lst:
            if i in text:
                for key in replacement_dict.keys():
                    text = text.replace(key,replacement_dict[key],1)
    return text

def parse_all_text_tags(xml_content):
    root = ET.fromstring(xml_content)

    text_tags = root.findall('.//Text')
    
    texts = [text.text.strip() for text in text_tags if text.text is not None]

    return texts

def replace_xml_content(xml_content, bus_data, new_bus_data):
    for i in range(len(bus_data)):
        xml_content = xml_content.replace(str(bus_data[i][0]), (str(new_bus_data[i][0])),1)
    return xml_content

def partition(array, low, high):
 
    pivot = array[high]
 
    i = low - 1
 
    for j in range(low, high):
        if array[j] <= pivot:
            i = i + 1
            (array[i], array[j]) = (array[j], array[i])
 
    (array[i + 1], array[high]) = (array[high], array[i + 1])
 
    return i + 1
 
 
def quicksort(array, low, high):
    if low < high:
        pi = partition(array, low, high)
 
        quicksort(array, low, pi - 1)
 
        quicksort(array, pi + 1, high)

def unallocated_number(allocated_dict,numbers):
    for i in numbers:
        if(allocated_dict[i] == 0):
            return i
    return -1

def bus_xml_change(source, dir_text):
    file_path = source

    try:
        with open(file_path, 'r',encoding = "UTF-8") as file:
            xml_content = file.read()
            all_texts = parse_all_text_tags(xml_content)

            numbers = []
            data = []
            bus_data = []
            for text in all_texts:
                if "Bus[" in text and not "HWBus[" in text:
                    lst = re.findall(r'\((.*?)\)', text)
                    for s in lst:
                        result = s.split(',')
                        data.append(result)
                        for s1 in result:
                            if "].Obj" in s1:
                                num = re.findall(r'\[(.*?)\]', s1)
                                n = (int(num[0]))
                                numbers.append(n)
                                bus_data.append([s1])

            quicksort(numbers, 0, len(numbers) - 1)
            numbers = list(np.unique(numbers))
            bus_num_list = []
            for i in numbers:
                bus_num_list.append([i])
            
            os.chdir(dir_text)
            file_name = os.getcwd()+"\\bus_list_numbers.csv"
            print("Warning if the file "+file_name+" is open in Microsoft Excel or any spreadsheet application please close it.\n")
            try:
                with open(file_name, mode='w', newline='',encoding='utf-8') as file:
                    writer = csv.writer(file)

                    writer.writerows([['Bus Number','Replacement Bus Number']])
                    writer.writerows(bus_num_list)
                    print("The file "+file_name+" is created")
            except FileNotFoundError:
                print(f"Error: File '{file_name}' not found.")
            except:
                print("The file was either not found or was being accessed through a spreadsheet software like Microsoft Excel")
        
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")


def bus_xml_replacement(source, dest, dir_text):
    file_path = source
    os.chdir(dir_text)
    f1 = os.getcwd()+"\\bus_list_numbers.csv"

    try:
        with open(file_path, 'r',encoding = "UTF-8") as file:
            xml_content = file.read()
            all_texts = parse_all_text_tags(xml_content)
            file_name = f1
            df = pd.read_csv(file_name)
            df = df.dropna()
            num = list(df['Bus Number'])
            new_num = list(df['Replacement Bus Number'])
            replacement_dict = {}
            for i in range(len(num)):
                replacement_dict[str(num[i])] = str(int(new_num[i]))
            root = parse_all_text_tags_root(xml_content)
            for text_tag in root.findall('.//Text'):
                original_text = text_tag.text.strip() if text_tag.text is not None else ''
                modified_text = replace_tags(original_text, replacement_dict)
                xml_content = xml_content.replace(original_text,modified_text)
            try:
                with open(dest, 'w',encoding = "UTF-8") as file:
                    file.write(xml_content)
                    print("Change has happened in ",dest)
            except FileNotFoundError:
                print(f"Error: File '{dest}' not found.")
        
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")

def list_bus(source,dir_text):
    file_path = source

    try:
        with open(file_path, 'r',encoding = "UTF-8") as file:
            xml_content = file.read()
            all_texts = parse_all_text_tags(xml_content)
            numbers = []
            data = []
            bus_data = []
            for text in all_texts:
                if "Bus[" in text and not "HWBus[" in text:
                    lst = re.findall(r'\((.*?)\)', text)
                    for s in lst:
                        result = s.split(',')
                        data.append(result)
                        for s1 in result:
                            if "Bus[" in s1:
                                num = re.findall(r'\[(.*?)\]', s1)
                                n = (int(num[0]))
                                numbers.append(n)
                                bus_data.append([s1])

            quicksort(numbers, 0, len(numbers) - 1)
            numbers = list(np.unique(numbers))
            bus_num_list = []
            for i in numbers:
                bus_num_list.append([i])
            bus_item_list = []
            for i in bus_data:
                bus_item_list.append(i[0])
            bus_item_list = list(set(bus_item_list))
            pattern = r'Bus\[\d+\]\.Obj'

            bus_item_list = [string for string in bus_item_list if re.search(pattern, string)]

            c = []
            for i in bus_item_list:
                count = 0
                lst = []
                for j in data:
                    for k in j:
                        if i == k:
                            count+=1
                            lst.append(j[0])
                c.append([i,count,lst])
            os.chdir(dir_text)
            f1 = os.getcwd()+"\\bus_count_with_tags.csv"
            print("Warning if the file "+f1+" is open in Microsoft Excel or any spreadsheet application please close it.\n")
            try:
                with open(f1, mode='w', newline='',encoding='utf-8') as file:
                    writer = csv.writer(file)

                    writer.writerows([['Bus Tags','Count','Original Tags','Replace Tags']])
                    writer.writerows(c)
                    print("The file "+f1+" is created")
            except FileNotFoundError:
                print(f"Error: File '{f1}' not found.")
            except:
                print("The file was either not found or was being accessed through a spreadsheet software like Microsoft Excel")
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")

def replace_bus_tags(source, dest, dir_text,start_number,end_number):
    file_path = source
    os.chdir(dir_text)
    f1 = os.getcwd()+"\\bus_count_with_tags.csv"
    csv_file_path = f1
    df = pd.read_csv(csv_file_path)
    try:
        with open(file_path, 'r', encoding="UTF-8") as file:
            xml_content = file.read()
            root = parse_all_text_tags_root(xml_content)
            all_texts = parse_all_text_tags_text(xml_content)
            bus_list = list(df['Bus Tags'])
            t = list(df['Original Tags'])
            rt = list(df['Replace Tags'])
            og_t = []
            rep_t = []
            tag_list = []
            count = list(df["Count"])
            replacement_dict = {}
            pattern = r'Bus\[(\d+)\].Obj'
            allocated_numbers = [int(re.search(pattern, s).group(1)) for s in bus_list if re.search(pattern, s)]
            allocated_numbers.sort()
            allocated_dict = {}
            numbers = range(start_number,end_number+1)
            for i in numbers:
                if i in allocated_numbers:
                    allocated_dict[i] = 1
                else:
                    allocated_dict[i] = 0
            for i in range(len(t)):
                if(count[i] == 1):
                    match = re.search(r"'(.*?)'",t[i])
                    if match:
                        extracted_string = match.group(1)
                        og_t.append(extracted_string)
                    else:
                        print("No match found.")
                    
                    rep_t.append(rt[i])
                elif(count[i] > 1):
                    lst = t[i].split(",")
                    for j in range(len(lst)):
                        match = re.search(r"'(.*?)'",lst[j])
                        if match:
                            extracted_string = match.group(1)
                            og_t.append(extracted_string)
                            if(j > 0):
                                number = unallocated_number(allocated_dict,numbers)
                                if number == -1:
                                    print("No unallocated numbers found")
                                    return
                                allocated_dict[number] = 1
                                replacement_dict[bus_list[i]] = "Bus["+str(number)+"].Obj"
                                tag_list.append(extracted_string)
                        else:
                            print("No match found.")
    
            for text_tag in root.findall('.//Text'):
                original_text = text_tag.text.strip() if text_tag.text is not None else ''
                modified_text = replace_bus(original_text, replacement_dict,tag_list)
                xml_content = xml_content.replace(original_text,modified_text)

            modified_file_path = dest

            with open(modified_file_path, 'w',encoding = "UTF-8") as file:
                file.write(xml_content)
                print("Change has happened")
            
            print(f"XML file '{modified_file_path}' has been created with replaced tags.")

    except FileNotFoundError:
        print(f"Error: File '{file_path}' or '{csv_file_path}' not found.")

def replace_tags_xml(source,dest,dir_text):
    file_path = source
    os.chdir(dir_text)
    f1 = os.getcwd()+"\\bus_count_with_tags.csv"
    csv_file_path = f1
    df = pd.read_csv(csv_file_path)
    df = df.dropna()
    try:
        with open(file_path, 'r', encoding="UTF-8") as file:
            xml_content = file.read()
            root = parse_all_text_tags_root(xml_content)
            all_texts = parse_all_text_tags_text(xml_content)
            bus_list = list(df['Bus Tags'])
            t = list(df['Original Tags'])
            rt = list(df['Replace Tags'])
            og_t = []
            rep_t = []
            count = list(df["Count"])
            bus_list = bus_list.sort()
            for i in range(len(t)):
                if(count[i] == 1):
                    match = re.search(r"'(.*?)'",t[i])
                    if match:
                        extracted_string = match.group(1)
                        og_t.append(extracted_string)
                    else:
                        print("No match found.")
                    
                    rep_t.append(rt[i])

            replacement_dict = {}
            for i in range(len(og_t)):
                replacement_dict[og_t[i]] = rep_t[i]

            for text_tag in root.findall('.//Text'):
                original_text = text_tag.text.strip() if text_tag.text is not None else ''
                modified_text = replace_tags(original_text, replacement_dict)
                xml_content = xml_content.replace(original_text,modified_text)

            modified_file_path = dest

            with open(modified_file_path, 'w',encoding = "UTF-8") as file:
                file.write(xml_content)
                print("Change has happened at "+modified_file_path)
            
            print(f"XML file '{modified_file_path}' has been created with replaced tags.")

    except FileNotFoundError:
        print(f"Error: File '{file_path}' or '{csv_file_path}' not found.")

def printOptionsText():
    return [
        "To list the tags with respective bus tags to either replace tags or to remove duplicate buses in csv",
        "To replace tags", 
        "To remove duplicate buses and replace them with the respective numbers available in the array",
        "To create the bus list to change bus names",
        "To replace bus numbers in the list"
    ]


@Gooey(program_name="XML Manipulation Tool",default_size = (1200, 700))
def main():
    parser = GooeyParser(description="Manipulating Rockwell XML files")
    lst = printOptionsText()
    
    parser.add_argument(
        "source", 
        widget="FileChooser",
        help = "Enter source L5X/L5K file path",
        metavar = "Source File"
        )
    
    parser.add_argument(
        "dest", 
        widget="FileChooser",
        help = "Enter destination L5X/L5K file path",
        metavar = "Destination File"
        )
    

    parser.add_argument(
        "dir_text",
        widget = "DirChooser",
        help = "Enter the folder path where you want to save all the csv files",
        metavar = "CSV files folder path"
        )
    
    parser.add_argument(
        "option", 
        choices=lst, 
        metavar="Function List", 
        help="Choose the function you want to execute"
        )
    
    parser.add_argument(
        "--start_number", 
        action = "store",
        help = "Enter start number of the range that the Buses are allocated\nIt is necessary if you want to call the function to eliminate duplicate buses",
        metavar = "Start Number",
        )
    
    
    parser.add_argument(
        "--end_number", 
        action = "store",
        help = "Enter end number of the range that the Buses are allocated\nIt is necessary if you want to call the function to eliminate duplicate buses",
        metavar = "End Number",
        )
    
    args = parser.parse_args()
    
    source = args.source
    dest = args.dest
    dir_text = args.dir_text
    ch = args.option
    start_number = args.start_number
    end_number = args.end_number

    
    if ch == lst[0]:
        list_bus(source,dir_text)
    elif ch == lst[1]:
        replace_tags_xml(source,dest,dir_text)
    elif ch == lst[2]:
        try:
            replace_bus_tags(source,dest,dir_text,int(start_number),int(end_number))
        except ValueError or TypeError:
            print("Incorrect Input in case of start and end numbers\nGo to Edit and Type the correct values")
        except TypeError:
            print("Blank Input\nGo to Edit and the correct values")
    elif ch == lst[3]:
        bus_xml_change(source,dir_text)
    elif ch == lst[4]:
        bus_xml_replacement(source,dest,dir_text)


if __name__ == "__main__":
    main()