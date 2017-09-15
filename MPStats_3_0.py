#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May 16 11:09:32 2017

@author: simonmartin
"""
import pandas as pd # Pandas use here for easy manipulation of data files
import numpy as np

def Geometry(nmodules,rowmax=6):
    """
    calculates plot grid dimensions – limited to rowmax rows
    returns tuple of nrows,ncols needed to fit all histograms (there may be blank spaces)
    uses integer division to work out number of columns
    """
    if nmodules<=rowmax:
        return (nmodules,1)#rows,columns
    else:
        # aim to make a grid that looks as full as possible
        # 
        # number of columns will be (nfiles/rowmax)+1(if (nfiles%rowmax)>=1)
        # number of rows: (nfiles/ncolumns)+1(if (nfiles%rowmax)>=1)
        ncolumns=nmodules/rowmax
        if (nmodules%rowmax)>=1:
            ncolumns=ncolumns+1
        nrows=nmodules/ncolumns
        if (nmodules%ncolumns)>=1:
            nrows=nrows+1
        return (nrows,ncolumns)

def StatFrame(filename='Module Marks.xls',programme_list=None,part_list=None,module_list=None,assess=False,debug=False,include_zeroes=False,dp=1,latex=False):
    """Creates dataframe of the marks for a module or set of modules
    Inputs:
        filename is an excelfile from LUSI system includes Module marks and programmes
        module_code is the code for the module of interest e.g. '15MPA201'
        programme_code (optional) returns data just for given programme e.g. MPUB01
        invalid codes will result in (nan,nan,0)
        assess=True will output average of coursework and exam marks
        dp: number of decimal places for results
        latex=True will output latex code for the result
        """
    if module_list==None: 
        module_list=[]
    if programme_list==None: 
        programme_list=[]
    if part_list==None:part_list=[]
# Method: build up dataframe of matching modules, then output this dataframe using df.to_latex()     
    # Create dataframe to put the required data in.
        # this dataframe will have at least len(programme_list)+3 columns
    #build a columns list
    col_list=['Module','All cohorts','StDev']
    if assess==True: col_list=col_list+['CW ave','Ex ave']
    if (programme_list != []):
        col_list=col_list+programme_list
    if debug==True: print('col_list='+str(col_list[:]))
    # read in the LUSI spreadsheet
    df=pd.read_excel(filename,skiprows=[0,1]) # have the LUSI file in df
    # some of the column names (may) have spaces. This is a problem in Pandas. Replace space with underscore
    cols = df.columns
    cols = cols.map(lambda x: x.replace(' ', '_') if isinstance(x, (str)) else x) # code based on: https://github.com/pandas-dev/pandas/issues/6508
    df.columns = cols
    if debug==True : print (df.columns.tolist())
    #if debug==True : print(df.head())
    # check to see if module list has values, otherwise go over all modules
    if (module_list==[]):
        # Now pull out list of modules – search "Module Code" column for unique values.
        module_list=df.Module_Code.unique()
    if debug==True : print ((module_list) )
    if debug==True : print (len(module_list) )
    datadf = pd.DataFrame(index=np.arange(0, len(module_list)), columns=col_list) # define size of dataframe for storing results
    for idx,module in enumerate(module_list):
        if debug==True : print (idx,module)
        # build up a list with the required info then append to the dataframe
        line=[module]
        if include_zeroes==False:
            #line.append(df["Module_Mark"][(df["Module_Code"] == module)][(df["Module_Mark"]!=0)].mean())
            line.append(np.round(df["Module_Mark"][(df["Module_Code"] == module)][(df["Module_Mark"]!=0)].mean(),1))
            line.append(np.round(df["Module_Mark"][(df["Module_Code"] == module)][(df["Module_Mark"]!=0)].std(),1))
            if assess==True:
                line.append(np.round(df["Cswk_Mark"][(df["Module_Code"] == module)][(df["Cswk_Mark"]!=0)].mean(),1))
                line.append(np.round(df["Exam_Mark"][(df["Module_Code"] == module)][(df["Exam_Mark"]!=0)].mean(),1)) # exclude zeroes
            for programme in programme_list:
                line.append(np.round(df["Module_Mark"][(df["Module_Code"] == module)& (df["Programme_Code"] == programme)][(df["Module_Mark"]!=0)].mean(),1))
        else:
            line.append(np.round(df["Module_Mark"][(df["Module_Code"] == module)].mean(),1))
            line.append(np.round(df["Module_Mark"][(df["Module_Code"] == module)].std(),1))
            if assess==True:
                line.append(np.round(df["Cswk_Mark"][(df["Module_Code"] == module)].mean(),1))
                line.append(np.round(df["Exam_Mark"][(df["Module_Code"] == module)].mean(),1))
        #datadf.loc[idx:idx,'Module':]=line # put line into dataf
            for programme in programme_list:
                line.append(np.round(df["Module_Mark"][(df["Module_Code"] == module)& (df["Programme_Code"] == programme)].mean(),1))
        if debug==True : print (line)
        datadf.loc[idx:idx,'Module':]=line # put line into dataf    
    if latex==True: return(print(datadf.to_latex(na_rep='--',index=False)))
    return (datadf)
"""     
"""  

def ModuleDF(filename='Module Marks.xls',module_name='',elements=False,debug=False): 
    """returns a dataframe of the results of module_name
        Inputs: Lusi results filename"""
    if module_name=='':
        print('Module name must be supplied')
        return
    df=pd.read_excel(filename,skiprows=[0,1]) # have the LUSI file in df
    # some of the column names (may) have spaces. This is a problem in Pandas. Replace space with underscore
    cols = df.columns
    cols = cols.map(lambda x: x.replace(' ', '_') if isinstance(x, (str)) else x) # code based on: https://github.com/pandas-dev/pandas/issues/6508
    df.columns = cols
    if debug==True : print (df.columns.tolist())
    # get results for module_name
    if elements==False:
        ResultList=[]
        ResultList=df["Module_Mark"][(df["Module_Code"] == module_name)]
        if debug: print (ResultList)
        #datadf=pd.DataFrame(ResultList)
        ResultFrame=pd.DataFrame({'Module':df["Module_Mark"][(df["Module_Code"] == module_name)]})
        if debug: print(ResultFrame)       
    else:
        ResultFrame=pd.DataFrame({'Module': df["Module_Mark"][(df["Module_Code"] == module_name)],'Cswk': df["Cswk_Mark"][(df["Module_Code"] == module_name)],'Exam':df["Exam_Mark"][(df["Module_Code"] == module_name)]})       
    return (ResultFrame)
    
    

def HistoModule(filename='Module Marks.xls',module_name='',bins=20,elements=False,debug=False,font_size=None):
    """creates a pandas histogram of the results for specified module
        Inputs: 
            filename is an excelfile from LUSI system includes Module marks and programmes
        module_name is the code for the module of interest e.g. '15MPA201'
        bins=number of bins for the data
        elements=True causes the results for both exam and CW to be displayed 
        """
    # method
    # get results for given module
    # display as histogram
    #
    # get dataframe for module_name
    DataF=ModuleDF(filename,module_name,elements,debug)
    # get results for module_name
    if elements==False:
        if debug: print(DataF)
        DataF.plot.hist(alpha=0.5,fontsize=font_size,range=(0,100),bins=bins,color='b')
    else:
        DataF.plot.hist(alpha=0.25,fontsize=font_size,bins=bins,stacked=False,sort_columns=False,color=['r','g','b'],range=(0,100))
    return

def HistoArray(filename='Module Marks.xls',module_list=[],bins=20,elements=False,debug=False):
    """creates a pandas histogram of the results for specified module
        Inputs: 
            filename is an excelfile from LUSI system includes Module marks and programmes
        module_list is a list of the modules of interest e.g. '15MPA201'
        bins=number of bins for the data
        elements=True causes the results for both exam and CW to be displayed 
        """
        
    """need to sort out code to follow list of modules"""
    
    # method
    # read in data
    # get results for given modules
    # display as array of histograms
    #
    # read in the LUSI spreadsheet
    df=pd.read_excel(filename,skiprows=[0,1]) # have the LUSI file in df
    # some of the column names (may) have spaces. This is sometimes a problem in Pandas. Replace space with underscore
    cols = df.columns
    cols = cols.map(lambda x: x.replace(' ', '_') if isinstance(x, (str)) else x) # code based on: https://github.com/pandas-dev/pandas/issues/6508
    df.columns = cols
    if debug==True : print (df.columns.tolist())
    #if debug==True : print(df.head())
    # check to see if module list has values, otherwise go over all modules
    if (module_list==[]):
        # Now pull out list of modules – search "Module Code" column for unique values.
        module_list=list(df.Module_Code.unique()) # see: https://chrisalbon.com/python/pandas_find_unique_values.html
    if debug==True : print ((module_list) )
    if debug==True : print (len(module_list) )
    # now work out dimensions of array
    #f, axarr = plt.subplots(nrows=nRows,ncols=nCols, sharex=True) # setup an array in which to put the plots
    #build a dataframe of result data frames then plot array using by keyword
    df2=df.loc[df['Module_Code'].isin(module_list)]
    print (df2)
    #for idx,module in enumerate(module_list):
        #df2=df2.append({'Module':df["Module_Mark"][(df["Module_Code"] == module)]})
    df2['Module_Mark'].hist(bins=bins,by=df['Module_Code'])
    return

def ProgPartHist(filename='Module Marks.xls',programme_list=[],part_list=[],module_list=[],bins=20,elements=False,debug=False):
    """ creates array of histograms of modules for given programmes/parts/modules
        if the lists are blank then it will go through everything"""
    # read in the LUSI spreadsheet
    df=pd.read_excel(filename,skiprows=[0,1]) # have the LUSI file in df
    # some of the column names (may) have spaces. This is sometimes a problem in Pandas. Replace space with underscore
    cols = df.columns
    cols = cols.map(lambda x: x.replace(' ', '_') if isinstance(x, (str)) else x) # code based on: https://github.com/pandas-dev/pandas/issues/6508
    df.columns = cols
    # check to see if programme list or part have values, of not set to all
    if (programme_list==[] and part_list==[]):
        # assume that user wants all programmes and parts
        programme_list=list(df.Programme_Code.unique())
        part_list=list(df.Part.unique())
        if (debug==True): print (part_list)
    if (module_list==[]): module_list=list(df.Module_Code.unique())# list of all modules as default
    #generate required array(s) and display
    if (programme_list==[]): # no programme list so go through producing summaries of each part in part_list
        for part in part_list:
            df2=df.loc[df['Part']==part]
            df3=df2.loc[df2['Module_Code'].isin(module_list)]
            df3['Module_Mark'].hist(bins=bins,by=df3['Module_Code'],range=(0,100))
    else: # have a list of programmes to go through
        if (part_list==[]):part_list=list(df.Part.unique())
        for programme in programme_list:
            print('prog=',programme)
            df2=df.loc[df['Programme_Code']==programme]
            for part in part_list:
                df3=df2.loc[df['Part']==part]
                df4=df3.loc[df3['Module_Code'].isin(module_list)]
                if not df4.empty:
                    df4['Module_Mark'].hist(bins=bins,by=df4['Module_Code'],range=(0,100))
                else:
                    print('No data to histogram. Check input lists')
    return

def ProgPartStats(filename='Module Marks.xls',programme_list=[],part=None,module_list=[],bins=20,elements=False,debug=False,assess=False,include_zeroes=False,dp=1,latex=False):
    """ creates table of module statistics for given programmes/parts/modules
        if the lists are blank then it will go through everything"""
    if (part==None):return('Error: must define part')
    # Method: build up dataframe of matching modules, then output this dataframe using df.to_latex()     
    # Create dataframe to put the required data in.
        # this dataframe will have at least len(programme_list)+3 columns
    #build a columns list
    col_list=['Module','All cohorts','StDev']
    if assess==True: col_list=col_list+['CW ave','Ex ave']
    if (programme_list != []):
        col_list=col_list+programme_list
    if debug==True: print('col_list='+str(col_list[:]))
    # read in the LUSI spreadsheet
    df=pd.read_excel(filename,skiprows=[0,1]) # have the LUSI file in df
    # some of the column names (may) have spaces. This is sometimes a problem in Pandas. Replace space with underscore
    cols = df.columns
    cols = cols.map(lambda x: x.replace(' ', '_') if isinstance(x, (str)) else x) # code based on: https://github.com/pandas-dev/pandas/issues/6508
    df.columns = cols
    df2=df.loc[df['Part']==part]#reduce dataset to only cover part of interest
    if (module_list==[]): module_list=list(df2.Module_Code.unique())# list of all modules as default
    #generate required array(s) and display
    df3=df2.loc[df2['Module_Code'].isin(module_list)]
        # Now have the modules, can make table
    module_list=df3.Module_Code.unique()
    if debug==True : print ((module_list) )
    if debug==True : print (len(module_list) )
    datadf = pd.DataFrame(index=np.arange(0, len(module_list)), columns=col_list) # define size of dataframe for storing results
    for idx,module in enumerate(module_list):
        if debug==True : print (idx,module)
        # build up a list with the required info then append to the dataframe
        line=[module]
        if include_zeroes==False:
            #line.append(df["Module_Mark"][(df["Module_Code"] == module)][(df["Module_Mark"]!=0)].mean())
            line.append(np.round(df3["Module_Mark"][(df3["Module_Code"] == module)][(df3["Module_Mark"]!=0)].mean(),dp))
            line.append(np.round(df3["Module_Mark"][(df3["Module_Code"] == module)][(df3["Module_Mark"]!=0)].std(),1))
            if assess==True:
                line.append(np.round(df3["Cswk_Mark"][(df3["Module_Code"] == module)][(df3["Cswk_Mark"]!=0)].mean(),1))
                line.append(np.round(df3["Exam_Mark"][(df3["Module_Code"] == module)][(df3["Exam_Mark"]!=0)].mean(),1)) # exclude zeroes
        # NOw add programme results (if requested)
            for programme in programme_list:
                line.append(np.round(df3["Module_Mark"][(df3["Module_Code"] == module)& (df3["Programme_Code"] == programme)][(df3["Module_Mark"]!=0)].mean(),1))
        else:
            line.append(np.round(df3["Module_Mark"][(df3["Module_Code"] == module)].mean(),1))
            line.append(np.round(df3["Module_Mark"][(df3["Module_Code"] == module)].std(),1))
            if assess==True:
                line.append(np.round(df3["Cswk_Mark"][(df3["Module_Code"] == module)].mean(),1))
                line.append(np.round(df3["Exam_Mark"][(df3["Module_Code"] == module)].mean(),1))
        #datadf.loc[idx:idx,'Module':]=line # put line into dataf
            for programme in programme_list:
                line.append(np.round(df3["Module_Mark"][(df3["Module_Code"] == module)& (df3["Programme_Code"] == programme)].mean(),1))
        if debug==True : print (line)
        datadf.loc[idx:idx,'Module':]=line # put line into dataf    
    if latex==True: return(print(datadf.to_latex(na_rep='--',index=False)))
    return (datadf)
   
def yearTable(filename='5years.xlsx',SheetName='PartA'):
    """Outputs latex code of table of 5 years of results for a given part's
    module results
    Input:filename is the excel file with the data in, 
    sheetname contains the data for the part to be tablularised"""
    xl=pd.ExcelFile(filename)
    df=xl.parse(SheetName)
    df2=df.round(1)
    #print(df.head())
    df2.set_index('Module Code', inplace=True)
    #print(df.head())
    print (df2.to_latex(column_format='clccccc',na_rep='--'))
    return
   





