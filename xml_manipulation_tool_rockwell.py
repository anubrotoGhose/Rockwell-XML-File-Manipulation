"""
XML Manipulation Tool for Rockwell PLC Programs

This Python script provides a set of functionalities to manipulate XML files commonly used in Rockwell Programmable Logic Controller (PLC) programs. The tool is designed to perform tasks such as replacing specific tags, managing bus numbers, and creating CSV files for analysis.

The script uses the ElementTree library for XML parsing and provides a command-line interface with a menu-driven approach for user interaction. Additionally, the script incorporates functionality to create and manipulate CSV files for storing and retrieving bus-related information.

Instructions:
1. Run the script, providing the path to the source L5X/L5K file.
2. Choose from the menu options to perform specific tasks, such as replacing tags, managing bus numbers, or listing tags with associated bus information.
3. Follow on-screen prompts for additional inputs and configurations.
4. The modified XML file and associated CSV files will be generated based on the chosen operations.

Note: Please ensure that the necessary Python libraries, such as xml.etree.ElementTree, csv, re, numpy, pandas, and os, are installed before running the script.

Author: Anubroto Ghose
Date: 07/12/2023
"""
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


# Function to find numbers from a list of numbers that have not yet been allocated to a bus
def unallocated_number(allocated_dict,numbers):
    for i in numbers:
        if(allocated_dict[i] == 0):
            return i
    return -1

# Function to perform list bus numbers in XML to a csv file
def bus_xml_list(source, dir_text):
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
            any_key = input("Warning if the file at "+file_name+" is open in Microsoft Excel or any spreadsheet application please close it.\nIf the file is closed then hit any key to continue:\t")
            try:
                with open(file_name, mode='w', newline='',encoding='utf-8') as file:
                    writer = csv.writer(file)

                    writer.writerows([['Bus Number','Replacement Bus Number']])
                    writer.writerows(bus_num_list)
            except FileNotFoundError:
                print(f"Error: File '{file_name}' not found.")
        
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
                    print("Change has happened")
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
            any_key = input("Warning if the file at "+f1+" is open in Microsoft Excel or any spreadsheet application please close it.\nIf the file is closed then hit any key to continue:\t")
            try:
                with open(f1, mode='w', newline='',encoding='utf-8') as file:
                    writer = csv.writer(file)

                    writer.writerows([['Bus Tags','Count','Original Tags','Replace Tags']])
                    writer.writerows(c)
            except:
                print(f"Error: File '{f1}' not found.")
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")

# Function to replace duplicated bus tags with unallocated numbers in the XML content
def replace_bus_tags(source, dest, dir_text):
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
            start_number = int(input("Enter start number of the range that the Buses are allocated:\t"))
            end_number = int(input("Enter end number of the range that the Buses are allocated:\t"))
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
    
    try:
        with open(file_path, 'r', encoding="UTF-8") as file:
            xml_content = file.read()
            root = parse_all_text_tags_root(xml_content)
            all_texts = parse_all_text_tags_text(xml_content)
            os.chdir(dir_text)
            f1 = os.getcwd()+"\\bus_count_with_tags.csv"
            csv_file_path = f1
            df = pd.read_csv(csv_file_path)
            df = df.dropna()
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
                print("Change has happened")
            
            print(f"XML file '{modified_file_path}' has been created with replaced tags.")

    except FileNotFoundError:
        print(f"Error: File '{file_path}' or '{csv_file_path}' not found.")
    except:
        print(f"Error: File '{file_path}' or '{csv_file_path}' not found.")

# Function to print options for the main menu
def printOptions():
    os.system('cls')
    print("Enter 0 to list the tags with respective bus tags to either replace tags or to remove duplicate buses in csv.")
    print("Enter 1 to replace tags.")
    print("Enter 2 to remove duplicate buses and replace them with the respective numbers available in the array.")
    print("Enter 3 to create the bus list to change bus names.")
    print("Enter 4 to replace bus numbers in the list.")
    print("Enter any other number to quit.")

def dest_path(source):
    y_n = input("Do you want the destination L5X/L5K file to be as same as source path (Y/N):\t")
    if(y_n == 'y' or y_n == 'Y'):
        return source
    else:
        d = input("Enter destination L5X/L5K file path:\t")
        return d

def main():
    source = input("Enter source L5X/L5K file path:\t")
    dest = dest_path(source)
    dir_text = input("Enter the folder path where you want to save all the csv files:\t")
    while True:
        printOptions()
        ch = input("Enter choice\n")
        if ch == '0':
            list_bus(source,dir_text)
        elif ch == '1':
            replace_tags_xml(source,dest,dir_text)
        elif ch == '2':
            replace_bus_tags(source,dest,dir_text)
        elif ch == '3':
            bus_xml_list(source,dir_text)
        elif ch == '4':
            bus_xml_replacement(source,dest,dir_text)
        else:
            break
        y_n = input("Do you want to quit (Y/N):\t")
        if(y_n == 'y' or y_n == 'Y'):
            break
        y_n = input("Do you want to keep the file configurations and folder paths unchanged (Y/N):\t")
        if(y_n == 'y' or y_n == 'Y'):
            continue
        y_n = input("Do you want the source L5X/L5K file to be as same as before (Y/N):\t")
        if(y_n == 'y' or y_n == 'Y'):
            pass
        else:
            source = input("Enter new source L5X/L5K file path:\t")
        y_n = input("Do you want the destination L5X/L5K file to be as same as before (Y/N):\t")
        if(y_n == 'y' or y_n == 'Y'):
            pass
        else:
            dest = dest_path(source)
        y_n = input("Do you want the folder path where you want to save all the csv files same as before (Y/N):\t")
        if(y_n == 'y' or y_n == 'Y'):
            pass
        else:
            dest = input("Enter the new folder path where you want to save all the csv files same as before (Y/N):\t")
if __name__ == "__main__":
    main()