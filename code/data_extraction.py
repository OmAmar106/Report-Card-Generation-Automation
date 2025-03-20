
def getalldata(filename):
    import pandas as pd
    from collections import defaultdict
    df = pd.read_excel(filename,sheet_name=None)

    finaldata = defaultdict(list)
    name = {}

    count = 1

    set1 = set()

    dmap = {}

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

                while k3<df[i].shape[1] and not pd.isna(df[i].iloc[2,k3]):
                    if not pd.isna(df[i].iloc[j,k3]):
                        finaldata[student].append((count,df[i].iloc[2,k3].replace('\n',' '),df[i].iloc[1,k3],df[i].iloc[3,k3],df[i].iloc[j,k3]))
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
                    continue
                
                L7 = []
                # now L4 has the back subjects and L2 has the summer term subjects
                set1.clear()
                for subject in L4:
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
                    L7.append((count,sub,cred,grade))
                # if L4:
                #     print(L4)

                for i1 in set1:
                    if i1 not in dmap:
                        # dmap[i1] = [input("Enter Full form of"+i1),input("Enter Course Code: ")]
                        dmap[i1] = [i1,"MAL 101"]
                
                for sem,subject,cred,grade in L7:
                    finaldata[student].append((sem,dmap[subject][0],dmap[subject][1],cred,grade))
                
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
                        # dmap[i1] = [input("Enter Full form of"+i1),input("Enter Course Code: ")]
                        dmap[i1] = [i1,"MAL 101"]
                
                for sem,subject,cred,grade in L7:
                    finaldata[student].append((sem,dmap[subject][0],dmap[subject][1],cred,grade))
                


        count += 1

    return (finaldata,name)