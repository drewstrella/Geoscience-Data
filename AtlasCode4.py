    # -*- coding: utf-8 -*-
"""
    Created on Sat Jul  2 02:53:09 2022
    
    @author: drews
"""
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
# from xlrd import XLRDError

#Establishing window
root=tk.Tk()
root.geometry("550x400")
root.pack_propagate(False)
root.resizable(0,0)



#adding instructions
instructiontext=str(  "Use the 'Select A File' button to select the Excel file containing the raw data. \n"
                    + "The path of the file will appear in the 'Open File' window. \n"
                    + "If the Excel file has multiple sheets, select the raw data sheet from the dropdown menu. \n"
                    + "Select 'Render Corrected File' button.\n"
                    + "A new file is created to preserve the original file with the original data and corrected data.\n"
                    + "The new file name and location will be displayed.\n"
                    + "To render additional files repeat instructions.\n"
                    + "\n"
                    + "Make sure file type is .xlsx and that the file is not currently open in another window.\n"
                    + "Each step may take a few seconds to load.")
#instruction button
button_instructions = tk.Button(root, text="Instructions", command=lambda: instruction_box())
button_instructions.place(rely=0.01, relx=0.01)


#Frame in window where buttons and file info is located
file_Frame = tk.LabelFrame(root, text="Select File") #label for 2nd frame
file_Frame.place(height=150, width=550, rely=0.1, relx=0)

#Frame with new file info
corrected_Frame = tk.LabelFrame(root, text = "New File Created")
corrected_Frame.place(height=150, width=550, rely=.55, relx=0)



# button to select file
button1 = tk.Button(file_Frame, text="Select A File", command=lambda: file_dialog())
button1.place(relx=0.01)

#button to render new file
button2= tk.Button(corrected_Frame, text="Render Corrected File", command=lambda: correct_excel_data())
button2.place(relx=0.01)

#Label in file frame that starts as "no file selected"
label_select=ttk.Label(file_Frame, text="No file selected")
label_select.place(rely=0.2, relx=0)
label_file = ttk.Label(file_Frame, text="")
label_file.place(rely=0.35, relx=0)

#Label for new file, starts as blank and changes to filename
label_new_file = ttk.Label(corrected_Frame, text=" ")
label_new_file.place(rely=.40, relx=0)

#label for sheet menu
label_menu= ttk.Label(file_Frame, text="")
label_menu.place(rely=.6)

# Label for new file changes from blank to "New File Saved As"
label_new_file2 = ttk.Label(corrected_Frame, text=" ")
label_new_file2.place(rely=.20, relx=0)

#dropdown menu for sheet selection
stv = tk.StringVar() #defining stv as a changable variable

#sheet selection menu (no placement as it only appears if needed)
optmenu = ttk.Combobox(file_Frame, textvariable = stv, state="readonly")
# optmenu.place(rely=.650, relx=0)


def instruction_box():
     tk.messagebox.showinfo("Instructions", instructiontext)
     
def file_dialog():
    filename=filedialog.askopenfilename(initialdir="/", title="Select A File", filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    
    #checking for supported filetype (.xlsx)
    try:
        if filename[-4:] != "xlsx":
            raise ValueError("Unsupported File Type \n"
                             +"Please select '.xlsx' File")
           
    except Exception as e:
        tk.messagebox.showinfo("Information", e)
       
        return None
    xl = pd.ExcelFile(filename)
    res = len(xl.sheet_names)
        
           
    xl = pd.ExcelFile(filename).sheet_names
    label_select["text"]="File Selected:"
    label_file["text"] = filename
    if res == 1:
            
        optmenu['values'] = []
        optmenu.set('')
    else:
        label_menu['text']="Select Sheet"
        optmenu.place(rely=.80, relx=0)
        optmenu['values'] = xl
    label_new_file["text"]= []
    label_new_file2["text"]= []
def correct_excel_data():
    file_path = label_file["text"]
    try:
        
        filename = r"{}".format(file_path)
        
        #determining if file is .xlsx
        # if filename[-4:] != "xlsx":
            
        #     raise ValueError("Incorrect File Type")
                # tk.messagebox.showerror("Information", "File Selection Invalid")
                # return None   
        # df=pd.read_excel(excel_filename)
    
    # def browse_file():
    #     filename = filedialog.askopenfilename()
        # originaldatadf = pd.read_excel(filename)    
        
        
        
    #    filename='200507_Earls Fruit Samps copy.xlsx'
        
        # #If more than one sheet use this one
        
        # sheetname='220328_Tykot Pottery'
        # originaldatadf = pd.read_excel(filename, sheet_name = sheetname, header=None)
       
        
        
        xl = pd.ExcelFile(filename)
        res = len(xl.sheet_names)
        
        if res == 1:
            
            #If only one sheet use this one
            originaldatadf = pd.read_excel(filename,header=None)
        else:
            sheetname = optmenu.get()
            originaldatadf = pd.read_excel(filename, sheet_name = sheetname, header=None)

        #display options
        pd.options.display.max_rows = 200 #for pandas display
        pd.options.display.precision = 15 #for pandas display
        
        elementlist = ['H', 'He', 'Li', 'Be', 'B', 'C', 'N', 'O', 'F', 'Ne', 'Na', 'Mg', 'Al', 'Si', 'P', 'S',
        'Cl', 'Ar', 'K', 'Ca', 'Sc', 'Ti', 'V', 'Cr', 'Mn', 'Fe', 'Co', 'Ni', 'Cu', 'Zn', 'Ga', 'Ge', 'As',
        'Se', 'Br', 'Kr', 'Rb', 'Sr', 'Y', 'Zr', 'Nb', 'Mo', 'Tc', 'Ru', 'Rh', 'Pd', 'Ag', 'Cd',
        'In', 'Sn', 'Sb', 'Te', 'I', 'Xe', 'Cs', 'Ba', 'La', 'Ce', 'Pr', 'Nd', 'Pm', 'Sm', 'Eu',
        'Gd', 'Tb', 'Dy', 'Ho', 'Er', 'Tm', 'Yb', 'Lu', 'Hf', 'Ta', 'W', 'Re','Os', 'Ir', 'Pt',
        'Au', 'Hg', 'Tl', 'Pb', 'Bi', 'Po', 'At', 'Rn', 'Fr', 'Ra', 'Ac', 'Th', 'Pa', 'U', 'Np',
        'Pu', 'Am', 'Cm', 'Bk', 'Cf', 'Es', 'Fm', 'Md', 'No', 'Lr', 'Rf', 'Db', 'Rg', 'Bh', 'Hs',
        'Mt', 'Ds', 'Rg', 'Cn', 'Nh', 'Fl', 'Mc', 'Lv', 'Ts', 'Og']
        
        
        # #removing blank rows
        originaldatadf.dropna(how='all', inplace = True) 
        originaldatadf.reset_index(drop=True,inplace=True)
        datadf=originaldatadf.copy()
        
        
        
        #reassigning column names
        datadf.columns = range(datadf.columns.size)
        
        #combining element with mass to differentiate between same elements with different mass
        
        datadf['datadf.shape[1]']=list(range(len(datadf)))
        
        datadf["element"] = datadf[1].astype(str)+ "-" + datadf[2].astype(str)
        
        #reassigning column names
        datadf.columns = range(datadf.columns.size)
        
        #Removing any rows that do not have an element to isolate data from nondata
        datadf = datadf[datadf[1].isin(elementlist)]
        
        
        
        
        #Find number of elements tested in each run
        
        unique=datadf[datadf.shape[1]-1].unique().tolist()
        
        #subjects is the number of elements tested
        subjects=len(unique)
        
        
        #Saving the starting point for the data
        
        datastart=datadf.index[0]
        
        
        
        #Finding standard deviation because standard deviatoion is the last subset and has no units
        #the previous subset will be the mean
        #and the number of subsets preceeding mean will be the number of trials per subject ID
        
        #finding all instances of NaN to find starting point of standard deviation (sd)
        sddf = datadf[datadf.iloc[:, 6].isnull()]
        
        
        
        
        
        #Finding starting point of standard deviation
        sdstart=sddf.iloc[0,datadf.shape[1]-2]
        
        #Standard deviation is always the last subset of data
        #(sdstart-subjects)/subjects=number of trials
        
        trials=int(((sdstart-subjects-datastart)/subjects))
        
        #Now I have the number of trials, the number of samples(subjects) in each trial, and their positions
        
        
        
        
        
        # Find first sample ID
        ## Making the assumption that the first Sample ID is "Blank"
        #if you find the first one the distance from that to the first sample is consistent
        
        # Number of total trials
        total=int(len(datadf)/(subjects*(trials+2)))
        
        # finding first instance of "BLANK" in slice of dataframe so i dont search entire dataset
        slice_of_df = originaldatadf.iloc[0:50, 0:4]
        
        # #finding all instances of "blank to find starting point"
        blankdf = slice_of_df[slice_of_df.iloc[:, 0:4].isin(['blank','Blank','BLANK','Blank 1','blank 1','BLANK1','blank1','BLANK 1'])]
        # dropping empty rows and columns
        blankdf.dropna(axis=0, how='all', inplace = True)
        blankdf.dropna(axis=1, how='all', inplace = True)
        
        #starting point of Blank
        blankstart=blankdf.iloc[0].name
        blankstartcolumn=blankdf.columns.values.tolist()
        
        sampleidstartcolumn=blankstartcolumn[0]
        
        
        
        
        
        #List of Sample IDs
        
        ID=[]
        ID.append(originaldatadf.iloc[blankstart,blankstartcolumn[0]])
        
        #finds first sample ID then using total samples and trials it finds the next one and so on
        for n in range(total-1):
            x=originaldatadf.iloc[int(datadf.iloc[subjects*(trials+2)*(n+1)-1].name)+1+blankstart,sampleidstartcolumn]
        
            #     x=originaldatadf.iloc[blankstart+int(datadf.iloc[(subjects*(trials+2))*(n)].name -datadf.index[0].name),sampleidstartcolumn]
            ID.append(x)
        
    
        
        
        
        #List of Sample Dates
        #The date seems to always be one cell below the sample ID
        datelist=[]
        datelist.append(originaldatadf.iloc[blankstart+1,blankstartcolumn[0]:3])
        
        for n in range(total-1):
            d=originaldatadf.iloc[int(datadf.iloc[subjects*(trials+2)*(n+1)-1].name)+2+blankstart,sampleidstartcolumn:3]
            datelist.append(d)
            
            
            
            
           #removing any blank columns from datelist
        #multiple columns were initially collected due to some dates be seperated across multiple cells
        
        datedf=pd.DataFrame(datelist)
        # datedf['New']=datedf.iloc[:,:].sum(1)
        
        datedf.dropna(axis=1, how='all', inplace = True)
        
        #Save each date as a list in a list of lists
        date=datedf.iloc[:,0:4].values.tolist()
        # print(date[0:5]+"\n\n")
        
        #Changing it to a list of strings so each date isn't its own list
        # date[0]=''.join(date[0])
        date = [''.join(i) for i in date]
        # print(date[0:5])
         
        # Reassembling the corrected data
        # Assembling a dataframe section by section for thousands of
        # lines is inefficient so I will build lists for each column
        
        #making lists each header per row
        A_header_list=['Quantitative Analysis - Comprehensive Report',
                       'Sample ID:',
                       'Sample Date/Time:', 
                       'Mean Values',
                      '']
        B_header_list=['','','','','Analyte']
        C_header_list=['','','','','Mass']
        D_header_list=['','','','','Meas. Intens. Mean']
        E_header_list=['','','','','Conc. Mean']
        F_header_list=['','','','','Conc. SD']
       
        #defining empty column lists to fill
        colA=[]
        colB=[]
        colC=[]
        colD=[]
        colE=[]
        colF=[]
                #defining starting value for n for function
        n=0
        #function to generate list for each column
        
        def columngenerator(col_num,col_list,col_head):
            for n in range(total):
                col=col_num
                #v=0 is for the mean values
                #v=1 is for the sd values
                # m pulls the correct column for each value
                if col==0:
                    v=0
                    m=col
                elif col==1:
                    v=0
                    m=col
                    col_head[1]=ID[n]
                    col_head[2]=date[n]
                elif col==2:
                    v=0
                    m=col
                elif col==3:
                    v=0
                    m=col
                elif col==4:
                    v=0
                    m=5
                elif col==5:
                    v=1
                    m=5
                #x is a list of date specified from the starting and endpoints found earlier with m being the correct column for that data
                x=datadf.iloc[(subjects*(trials+2))*(n)+subjects*(trials+v):subjects*(trials+2)*(n)+(subjects*(trials+v))+subjects,m].tolist()
                
                col_list.extend(col_head)
                col_list.extend(x)
                
        #         return col_list
        n=0
        #using function to generate lists to make dataframe
        columngenerator(0,colA,A_header_list)
        columngenerator(1,colB,B_header_list)
        columngenerator(2,colC,C_header_list)
        columngenerator(3,colD,D_header_list)
        columngenerator(4,colE,E_header_list)
        columngenerator(5,colF,F_header_list)
        
        #zipping the lists to make a dataframe
        zipped = list(zip(colA,colB,colC,colD,colE,colF))
        
        #making dataframe
        corrected_df=pd.DataFrame(zipped)
        
        
        #length of header+data (total number of cells per corrected data trial)
        total_rows=len(colA)
        
        corrected_trial_len=int(total_rows/total)
        
        # saving columns e and f (4 and 5) as numeric values so I can perform opperations on them
        e= pd.to_numeric(corrected_df[4], errors='coerce')
        f= pd.to_numeric(corrected_df[5], errors='coerce')
        
        #Performing operation ((f/e)*100) to form new column in corrected_df
        colG= ((f/e)*100).tolist()
        
        #col_header_list for col 6 aka col G
        G_header_list=['','','','','RSD %']
        
        #function to replace header for columns for future data manipulation
        def replace_header(col_list,col_head):
            i=0
            while i<total_rows:
                repl_list_strt_idx = i
                repl_list_end_idx = i+5
                col_list[repl_list_strt_idx : repl_list_end_idx] = col_head
                i+= corrected_trial_len
            return col_list
        
        # print(colG)
        replace_header(colG,G_header_list)
        corrected_df[6]=colG        
        
        
        
        
        # Creating new file with original data and corrected data
        index = filename.find('.')
        corrected_filename = filename[:index] + '_Corrected ' + filename[index:]
        
        with pd.ExcelWriter(corrected_filename, engine='xlsxwriter') as writer:
            originaldatadf.to_excel(writer, sheet_name='Raw_Data', index=False, header=False)
            corrected_df.to_excel(writer, sheet_name='Corrected Data', index=False,header=False)
        
        label_new_file2["text"]= "New File Saved As"        
        label_new_file["text"]= corrected_filename   
             
    # except Exception as e:
    #  tk.messagebox.showinfo("Info", e)
    # except ValueError:
    #     tk.messagebox.showerror("Information", "File Selection Invalid")
    #     return None

    # except FileNotFoundError:
    #     tk.messagebox.showerror("Information", f"No such file as {file_path}")
    #     return None
    except Exception as e:
        tk.messagebox.showinfo("Information", e)
     
        return None



    # except PermissionError:
    #     tk.messagebox.showinfo("Error", "Permission Denied \n\n File can not be called while open. Close Excel file to continue.")
    #     return None
    
"""exceptions for not an excel file
as well as for Permission Denied File still open"""

root.mainloop()        
        