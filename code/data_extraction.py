
def getalldata(filename):
    import pandas as pd
    from collections import defaultdict
    df = pd.read_excel(filename,sheet_name=None)

    finaldata = defaultdict(list)
    name = {}

    count = 1

    for i in df:
        if i.upper()=='CGPA':
            continue
        for j in range(7,df[i].shape[0]):
            if not pd.isna(df[i].iloc[j,1]):
                student = (df[i].iloc[j,1].lower())
                # finaldata[student].append()
                if student not in name and not pd.isna(df[i].iloc[j,2]):
                    # print(df[i].iloc[j,2])
                    name[student] = df[i].iloc[j,2].title()
                
                k3 = 3

                while k3<df[i].shape[1] and not pd.isna(df[i].iloc[2,k3]):
                    if not pd.isna(df[i].iloc[j,k3]):
                        finaldata[student].append((count,df[i].iloc[2,k3].title(),df[i].iloc[1,k3],df[i].iloc[3,k3],df[i].iloc[j,k3]))
                    k3 += 1
            
        count += 1

    return (finaldata,name)