import os
import csv
import pandas as pd
from itertools import zip_longest

def get_file_list():
    '''
    Returns
    -------
    list
        list contains names of files found in folder. Alarms user if no files found.
    '''
    if len(os.listdir()) == 1:
        print(f"Working directory: {os.getcwd()}")
        print(f"Found the following files: {os.listdir()}\n")
    else:
        print(f"Working directory: {os.getcwd()}")
        print(f"Found the following files: {os.listdir()}\n")
    return os.listdir()

def check_count_of_files(list_files):
    '''
    Parameters
    ----------
    list_files : list
        list of files found in folder.
        
    Returns
    -------
    bool
        pass if number of files is ok to proceed further

    '''
    if len(list_files) < 3:
        print('Not enough total files in folder')
        return False
    else: 
        return True
        
def find_txt_file(list_files_folder):
    '''
    Parameters
    ----------
    list_files : list
        list of files found in folder.
        
    Returns
    -------
    list
        returns list of files where .txt is found
    '''
    txt_files_list = []
    for i in range(0,len(list_files_folder)):
        if ".txt" in list_files_folder[i]:
            txt_files_list.append(list_files_folder[i])
    return txt_files_list
    
def check_count_txt_files(list_txt_files):
    '''
    Parameters
    ----------
    list_files : list
        list of files with .txt found from find_txt_file()
        
    Returns
    -------
    bool
        pass if number of files is ok to proceed further
    '''
    if len(list_txt_files) == 2:
        return True
    else:
        print("Number of .txt files is not adequate")
        return False

def process_placement_list(file_name):
    '''
    Parameters
    ----------
    file_name : string
        name of placement file that needs to be processed.

    Returns
    -------
    number_list : list
        list of part numbers extracted from PL file.

    '''
    
    read_file = pd.read_csv(file_name,delimiter = "\t", header = None)
    read_file = read_file.drop([0], axis=1)
    read_file = read_file.drop(range(2,22), axis=1)
    
    long_number_list = []
    for n in range(0,len(read_file)):
        long_number_list.append(read_file[1][n])
    
    
    number_list  = [part.replace('System\\', '') for part in long_number_list]
    del long_number_list
    return number_list

def find_unique_pns(number_list1,number_list2):
    '''
    Parameters
    ----------
    number_list1 : list
        List of PNs from PL.
    number_list2 : list
        List of PNs from PL.

    Returns
    -------
    List
        List of unique PNs from PL - number_list1.

    '''
        
    inside_set1 = set(number_list1)
    inside_set2 = set(number_list2)
    
    if len(list(inside_set1.difference(inside_set2))) == 0:
        return ["No unique components"]
    else:
        return list(inside_set1.difference(inside_set2))

def common_pns(number_list1,number_list2):
    
    inside_set1 = set(number_list1)
    inside_set2 = set(number_list2)
    
    return list(inside_set1.intersection(inside_set2))

def generate_report(file_name1,file_name2,unique_list1,unique_list2,common_list):
    
    rep_list = [unique_list1,unique_list2,common_list]
    export_data = zip_longest(*rep_list, fillvalue = '')
    del rep_list
    
    with open(f'Results for PLs {file_name1} and {file_name2}.csv', 'w', encoding="ISO-8859-1", newline='') as myfile:
      wr = csv.writer(myfile)
      wr.writerow((f"Unique PNs in {file_name1}", f"Unique PNs in {file_name2}","Common PNs"))
      wr.writerows(export_data)  
      myfile.close()
    print("Report generated. Look in folder for .csv file")

if __name__ == "__main__":
    
    files_in_folder = get_file_list()
     
    
    if check_count_of_files(files_in_folder):
                
        txt_files_in_folder = find_txt_file(files_in_folder)
        
        if check_count_txt_files(txt_files_in_folder):
            
            txt_filename1 = txt_files_in_folder[0]
            txt_filename2 = txt_files_in_folder[1] # Useful for further processing
            
            list1 =process_placement_list(txt_filename1)
            list2 = process_placement_list(txt_filename2)
            
            unique1 = find_unique_pns(list1, list2)
            unique2 = find_unique_pns(list2, list1)
            common = common_pns(list1,list2)
            
            generate_report(txt_filename1, txt_filename2, unique1, unique2, common)
    else:
        pass

os.system("pause")




    
