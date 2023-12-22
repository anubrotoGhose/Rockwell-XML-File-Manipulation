"""
XML Manipulation Tool

This Python script provides a set of functionalities to manipulate XML files commonly used in Rockwell Programmable Logic Controller (PLC) programs. The tool is designed to perform tasks such as replacing specific tags, managing bus numbers, and creating CSV files for analysis.

The script uses the ElementTree library for XML parsing and provides a GUI interface and interactionx. Additionally, the script incorporates functionality to create and manipulate CSV files for storing and retrieving bus-related information.

Dependencies:
- gooey
- xml.etree.ElementTree
- csv
- re
- numpy
- pandas
- os

Usage:
1. Run the script and choose the desired function through a graphical user interface (GUI).
2. Provide the source L5X/L5K file path, destination file path, and the folder path to save CSV files.
3. Optionally, specify start and end numbers for functions that eliminate duplicate buses.

Note: Make sure to close any open spreadsheet applications like Microsoft Excel before creating or modifying files.

Author: Anubroto Ghose
Date: 07/12/2023
"""

from gooey import Gooey, GooeyParser

import xml.etree.ElementTree as ET
import csv
import re
import numpy as np
import pandas as pd
import os

# Function to parse all text tags at the root level of XML
def parse_all_text_tags_root(xml_content):
    root = ET.fromstring(xml_content)
    return root

# Function to parse all text tags within Text elements of XML
def parse_all_text_tags_text(xml_content):
    root = ET.fromstring(xml_content)

    text_tags = root.findall('.//Text')
    
    texts = [text.text.strip() for text in text_tags if text.text is not None]

    return texts

# Function to replace specific tags in a text using a replacement dictionary
def replace_tags(text, replacement_dict):
    pattern = r'Bus\[\d+\]\.Obj'
    if re.search(pattern, text):
        for key in replacement_dict.keys():
            text = text.replace(key,replacement_dict[key],1)
    return text

# Function to replace specific bus-related tags in a text using a replacement dictionary and a list of bus items
def replace_bus(text, replacement_dict,lst):
    pattern = r'Bus\[\d+\]\.Obj'
    if re.search(pattern, text):
        for i in lst:
            if i in text:
                for key in replacement_dict.keys():
                    text = text.replace(key,replacement_dict[key],1)
    return text

# Function to parse all text tags within Text elements of XML
def parse_all_text_tags(xml_content):
    root = ET.fromstring(xml_content)

    text_tags = root.findall('.//Text')
    
    texts = [text.text.strip() for text in text_tags if text.text is not None]

    return texts

# Quicksort algorithm for sorting
# Partition
def partition(array, low, high):
 
    pivot = array[high]
 
    i = low - 1
 
    for j in range(low, high):
        if array[j] <= pivot:
            i = i + 1
            (array[i], array[j]) = (array[j], array[i])
 
    (array[i + 1], array[high]) = (array[high], array[i + 1])
 
    return i + 1
 
# Quicksort
def quicksort(array, low, high):
    if low < high:
        pi = partition(array, low, high)
 
        quicksort(array, low, pi - 1)
 
        quicksort(array, pi + 1, high)

# Function to find numbers from a list of numbers that have not yet been allocated to a bus
def unallocated_number(allocated_dict,numbers):
    for i in numbers:
        if(allocated_dict[i] == 0):
            return i
    return -1

# Function to perform list bus numbers in XML to a csv file
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
                print("Error: The files were either not found or "+file_name +" was being accessed through a spreadsheet software like Microsoft Excel")
        
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")

# Function to replace bus numbers in the XML content
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
            except :
                print(f"Error: File '{dest}' not found.")
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")

# Function to list all the bus numbers with their count and respective tags
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
                print("Error: The file was either not found or was being accessed through a spreadsheet software like Microsoft Excel")
    except FileNotFoundError:
        print(f"Error: File '{file_path}' or '{dir_text}' not found.")

# Function to replace duplicated bus tags with unallocated numbers in the XML content
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

# Function to replace specific tags in the XML content
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
            try:
                with open(modified_file_path, 'w',encoding = "UTF-8") as file:
                    file.write(xml_content)
                    print(f"XML file '{modified_file_path}' has been created with replaced tags.")
            except:
                print("Error: File  "+modified_file_path+" not found")
    except FileNotFoundError:
        print(f"Error: File '{file_path}' or '{csv_file_path}' not found.")


#Function to find instances of a particular tag
def num_par_tag(source,dest,dir_text,label):
    file_path = source
    try:
        with open(file_path, 'r', encoding="UTF-8") as file:
            xml_content = file.read()
            os.chdir(dir_text)
            f1 = os.getcwd()+"\\bus_count_with_tags.csv"
            csv_file_path = f1
    
            df = pd.read_csv(csv_file_path)
            lst = list(df['Original Tags'])
            result_list = []

            for given_string in lst:
                extracted_strings = re.findall(r"'(.*?)'", given_string)
                result_list.extend(extracted_strings)
            
            tag_dash_list = []
            for i in result_list:
                dash_string = i.replace("_", "-")
                tag_dash_list.append(dash_string)
                        
            label_dict = {}
            for i in tag_dash_list:
                occurrences = xml_content.count(i)
                label_dict[i] = occurrences

            replace_replace_label_list = []
            replace_label_list = []
            for i in tag_dash_list:
                if i == label:
                    for j in range(label_dict[i]):
                        replace_label = "type" + str(int(j+1))
                        replace_label_list.append(replace_label)
                        replace_replace_label = label + replace_label
                        replace_replace_label_list.append(replace_replace_label)
                        xml_content = xml_content.replace(label,replace_label,1)
            
            for i in range(len(replace_label_list)):
                xml_content = xml_content.replace(replace_label_list[i],replace_replace_label_list[i],1)
                print(replace_replace_label_list[i])
            # print("Replace label")
            # print(replace_replace_label_list[0])
            modified_file_path = dest
            with open(modified_file_path, 'w',encoding = "UTF-8") as file:
                file.write(xml_content)
                print("Change has happened")
            print(f"XML file '{modified_file_path}' has been created with replaced tags.")

    except FileNotFoundError:
        print(f"Error: File '{file_path}' or '{csv_file_path}' not found.")
    except:
        print(f"Error: File '{file_path}' or '{csv_file_path}' not found.")

# To list all the comments under rung tags with their specific properties
def extract_comments(source, dir_text):
    file_path = source
    try:
        with open(file_path, 'r', encoding="UTF-8") as file:
            xml_content = file.read()
            os.chdir(dir_text)
            # Parse the XML content
            root = ET.fromstring(xml_content)
            tasks_dict = {}

            for task in root.findall('.//Task'):
                task_name = task.get('Name')
                scheduled_programs = [program.get('Name') for program in task.findall('.//ScheduledProgram')]
                
                tasks_dict[task_name] = scheduled_programs
            # print(tasks_dict)
            reverse_dict = {}

            for task, programs in tasks_dict.items():
                for program in programs:
                    reverse_dict[program] = task
            
            # print(reverse_dict)
            txt = ""
            txt_list = []
            c = 0
            k = 0
            for program in root.findall('.//Program'):
                program_name = program.get('Name')
                task_name = reverse_dict.get(program_name, "Program not found in any task")
                txt = txt + f"Program: {program_name}"+ "\n"

                for routine in program.findall('.//Routine'):
                    routine_name = routine.get('Name')
                    txt = txt + f"  Routine: {routine_name}"+"\n"

                    for rung in routine.findall('.//Rung'):
                        comment_element = rung.find('.//Comment/LocalizedComment')
                        if comment_element is not None:
                            comment_text = comment_element.text.strip()
                            c+=1
                            txt = txt + f"      Task Name: {task_name}, Routine Name: {routine_name},  Rung Number:{rung.get('Number')}, Rung Type: {rung.get('Type')}, Comment Lang: {comment_element.attrib.get('Lang')}:, Comment:\n            {comment_text}" + "\n"
                            lst = [task_name, program_name, routine_name,rung.get('Number'),rung.get('Type'),comment_element.attrib.get('Lang'),comment_text]
                            txt_list.append(lst)
                        else:
                            k+=1
                            txt = txt + f"      Task Name: {task_name}, Routine Name: {routine_name},  Rung Number:{rung.get('Number')}, Rung Type: {rung.get('Type')}, Comment Lang: No Language, Comment: No comment" + "\n"
                            lst = [task_name, program_name, routine_name, rung.get('Number'),rung.get('Type'),"No Language", "No Comment"]
                            txt_list.append(lst)
            
            modified_file_path = os.getcwd()+"\\extracted_comments_under_rungs.txt"
            print("Total Number of rungs which have a comment = ",c)
            print("Total Number of rungs which do not have a comment = ",k)
            print("Total Number of rungs = ",k+c)
            with open(modified_file_path, 'w',encoding = "UTF-8") as file:
                file.write(txt)
                print("Change has happened")
            print(f"XML file '{modified_file_path}' has been created with replaced tags.")



            modified_file_path = os.getcwd()+"\\extracted_comments_under_rungs.csv"
            with open(modified_file_path, 'w',encoding = "UTF-8") as file:
                writer = csv.writer(file)

                writer.writerows([["Task Name","Program Name", "Routine Name", "Rung Number", "Rung Type", "Language", "Comment"]])
                writer.writerows(txt_list)

                print("Change has happened")
            print(f"XML file '{modified_file_path}' has been created with replaced tags.")

    except:
        print("Error: The Source file or folder paths are not found")

# Function to List the functions
def printOptionsText():
    return [
        "List tags with corresponding bus object numbers for tag replacement or removal of duplicate buses in bus_count_with_tags.csv",
        "To replace tags", 
        "To remove duplicate buses and replace them with the respective numbers available in the array",
        "To find instances of a particular tag/label.",
        "To list all the comments under rung tags.",
        "List the bus numbers in the bus_list_numbers.csv file for facilitating bus object number changes",
        "To replace bus object numbers in the xml file"
    ]

# main program
@Gooey(program_name="XML Manipulation Tool",default_size=(1455, 630))
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
    

    parser.add_argument(
        "--label", 
        action = "store",
        help = "Enter the label to find its number of instances",
        metavar = "Label",
        )

    args = parser.parse_args()
    
    source = args.source
    dest = args.dest
    dir_text = args.dir_text
    ch = args.option
    start_number = args.start_number
    end_number = args.end_number
    label = args.label

    try:
        if ch == lst[0]:
            list_bus(source,dir_text)
        elif ch == lst[1]:
            replace_tags_xml(source,dest,dir_text)
        elif ch == lst[2]:
            try:
                replace_bus_tags(source,dest,dir_text,int(start_number),int(end_number))
            except ValueError or TypeError:
                print("Error: Incorrect Input in case of start and end numbers\nGo to Edit and Type the correct values")
            except TypeError:
                print("Error: Blank Input\nGo to Edit and fill the values of start and end number")
        elif ch == lst[3]:
            num_par_tag(source,dest,dir_text,label)
        elif ch == lst[4]:
            extract_comments(source,dir_text)
        elif ch == lst[5]:
            bus_xml_change(source,dir_text)
        elif ch == lst[6]:
            bus_xml_replacement(source,dest,dir_text)
    except:
        print("Error:\tEnter Correct File or Folder Paths")

if __name__ == "__main__":
    main()
