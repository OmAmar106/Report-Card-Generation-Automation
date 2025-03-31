import json
import os

def getalldata(filename):
    import pandas as pd
    from collections import defaultdict
    df = pd.read_excel(filename,sheet_name=None)

    finaldata = defaultdict(list)
    name = {}

    count = 1

    set1 = set()

    dmap = {}

    with open(os.path.dirname(os.path.abspath(__file__))+"\\data\\data.json", "r") as file:
        dmap = json.load(file)
    
    print("Enter y in case the name is already the full form")

    for i in df.keys():
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
                reexam = []

                coursesinback = {}

                set1.clear()

                while k3<df[i].shape[1] and not pd.isna(df[i].iloc[2,k3]):
                    if not pd.isna(df[i].iloc[j,k3]):
                        grd = df[i].iloc[j,k3]
                        if grd[-1]=="*":
                            grd = grd[:-1]
                            coursesinback[df[i].iloc[2,k3].replace('\n',' ').strip()] = df[i].iloc[1,k3]
                            set1.add(df[i].iloc[2,k3].replace('\n',' ').strip())
                            reexam.append((count,df[i].iloc[2,k3].replace('\n',' ').strip(),df[i].iloc[3,k3],grd+'*'))
                            grd = "FF"
                        finaldata[student].append((count,df[i].iloc[2,k3].replace('\n',' ').strip(),df[i].iloc[1,k3],df[i].iloc[3,k3],grd))
                    k3 += 1

                try:

                    back = (df[i].iloc[j,k3])
                    L2 = []
                    L4 = []
                    L3 = []
                    i1 = 0
                    while (i1<len(back)):
                        if back[i1]==')':
                            if back[i1+1:min(i1+8,len(back))]==" S.TERM" or back[i1+1:min(i1+7,len(back))]=="S.TERM":
                                L2.append(L3)
                                i1 += 7
                            else:
                                L4.append(L3)
                            L3 = []
                        else:
                            if not L3 and back[i1]==' ':
                                i1 += 1
                                continue
                            L3.append(back[i1])
                        i1 += 1
                except:
                    if reexam==[]:
                        continue
                
                L7 = []
                # now L4 has the back subjects and L2 has the summer term subjects
                
                

                for subject in L4:
                    arr = subject
                    try:
                        if arr[0]==",":
                            arr = arr[1:]
                    except:
                        continue
                    grade = arr[0]+arr[1]
                    if "W" in grade:
                        grade = "W"
                    flag = False
                    if 2<len(arr) and arr[2]=="*":
                        flag = True
                        grade += arr[2]
                    cred = (arr[-1])
                    sub = []
                    f = arr.index('(')+1

                    while arr[f]!=',' and arr[f]!="=":
                        sub.append(arr[f])
                        f += 1
                    if arr[f]=="=":
                        sub.pop()
                        sub.pop()
                        
                    sub = ''.join(sub)
                    set1.add(sub)
                    if not flag:
                        L7.append((count,sub,cred,grade))
                    else:
                        reexam.append((count,sub,cred,grade))
                # if L4:
                #     print(L4)

                for i1 in set1:
                    if i1 not in dmap:
                        # dmap[i1] = [input("Enter Full form of"+i1),input("Enter Course Code: ")]
                        if i1 not in coursesinback and i1 not in dmap:
                            sub = input("Enter Full form of "+i1+" : ")
                            if sub=="y":
                                sub = i1
                            dmap[i1] = [sub,input("Enter Course Code of "+i1+" : ")]
                        elif i1 not in dmap:
                            dmap[i1] = [i1,coursesinback[i1]]
                
                for sem,subject,cred,grade in L7:
                    finaldata[student].append((sem,dmap[subject][0],dmap[subject][1],cred,grade))

                for sem,subject,cred,grade in reexam:
                    finaldata[student].append((sem+0.25,dmap[subject][0],dmap[subject][1],cred,grade[:-1]))
                
                L7 = []


                set1.clear()
                for subject in L2:
                    arr = subject
                    try:
                        if arr[0]==",":
                            arr = arr[1:]
                    except:
                        continue
                    grade = arr[0]+arr[1]
                    if "W" in grade:
                        grade = "FF"
                    if 2<len(arr) and arr[2]=="*":
                        grade += arr[2]
                    cred = (arr[-1])
                    sub = []
                    f = arr.index('(')+1
                    while arr[f]!=',':
                        sub.append(arr[f])
                        f += 1
                    sub = ''.join(sub)
                    set1.add(sub)
                    L7.append((count+(0.5),sub,cred,grade))
                # if L4:
                #     print(L4)

                for i1 in set1:
                    if i1 not in dmap:
                        sub = input("Enter Full form of "+i1+" : ")
                        if sub=="y":
                            sub = i1
                        dmap[i1] = [sub,input("Enter Course Code of "+i1+" : ")]
                
                for sem,subject,cred,grade in L7:
                    finaldata[student].append((sem,dmap[subject][0],dmap[subject][1],cred,grade))
                


        count += 1
    # print(dmap)
    with open(os.path.dirname(os.path.abspath(__file__))+"\\data\\data.json", "w") as file:
        json.dump(dmap, file, indent=4)

    return (finaldata,name)
