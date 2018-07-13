def runner(OriginalCSVInput):
    import pandas as pd
    name = pd.read_csv(OriginalCSVInput)
    name
   
    csvname = []
    docxname = []
    Filename = []
    
    for index, row in name.iterrows():  

        csvname.append(row['Company Name'] + "_" + row['Position'] + '.csv')
        csvfile = row['Company Name'] + "_" + row['Position'] + '.csv'
        docxname = row['Company Name'] + "_" + row['Position'] + '_' + "CoverLetter" + '.docx'
        Filename.append((csvfile, docxname)) 
    
    return [csvname, Filename]
    

