# VERISK competition to identify the fradulant name, location and scheme
# Vivek Pandiyan
# Project start date: 11th December, 2015
# Phase 1 modification: Feb 25th, 2016 - Implementing Patterns andd database and improved performance 

#------------------------------------------------------------------
# Programe uses Class 3 modifiers in NER to identify names
# Creates a database of articles and Fraudulent names to improve the effeciency of the program
# Generates a Excel output 
# Performance tuned
#-----------------------------------------------------------------

#Library imported for identifying the pronouns
from nltk import tokenize
#Library imported for sentimental analysis
from vaderSentiment.vaderSentiment import sentiment as senty
#Library used for importing counter for identifying the frequency
from collections import Counter
#Library imported to import excel file
import xlrd
#Library imported to export excel file
import xlwt
#Library imported for identifying the Name of a person
import ner
#Library imported for spliting words based on a pattern
import re
#Library for writing the results in excel sheet
from xlutils.copy import copy
#Library imported for identifying the Location
import geocoder

#Main function 
def main():
    file_location = "C:/Users/Vivek/Desktop/Class Notes/Python/Project_PY/ISO/ISO.xlsx"
    workbook = xlrd.open_workbook(file_location)
    print "Enter the sheet number"
    try:
        w=int(raw_input()) -1
    except ValueError:
        print("That's not an integer input!")
    sheet = workbook.sheet_by_index(w)
    data = [[sheet.cell_value(r,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    n,m=getinput(len(data))
    wb = copy(workbook)
    s = wb.get_sheet(w)
    for e in range(len(data))[n:m]:
        print "-"*20 
        print "CASE NUMBER %r" %(e)
        print "-"*20 
        mydat = data[e][0]
        mydata = cleanrawdata(mydat)
        Fraudlentname,StateandCity,Scheme,Govofficials,Databaseflag = databasefile(mydat)        
        GovName = initialze()
        LocName = initialze()
        words = re.findall(r'\w+', mydata)
        criminalactivity,family= Fcheck(mydata)
        if criminalactivity == "True" and Databaseflag != "True":
            word_counts = Counter(words) # Count the words
            Name,Orgname,Noun,Suspect,Familynames,Realname,Aliasname = name(word_counts,mydata,family)
            State, City = location(Noun,mydata)
            if Name != ["No Person Name available"] or Suspect != []:
                try:
                    Lname = lastname(Name)
                    Fullname = fulname(Name,Lname)
                    title,names,Fullname,mydata = plural(mydata,Fullname)
                    crimetitle=list(set(re.findall('.*?(%s).*?'%vhatecrime,mydata,re.IGNORECASE)))
                    crimetitle+=list(set(re.findall('.*?(%s).*?'%capcrime,mydata)))
                    govtitle, govname,mydata = findname(Fullname,mydata,crimetitle,words,Lname)
                    FraudulentName = list(set(Name) -set(govname)) 
                    govtitle += title;govname.append(names)
                    if govtitle != []:
                        for i in range(len(govtitle)):
                            print govtitle[i] + ":" + str(govname[i]) +"\n" 
                            GovName.append(govtitle[i] + " : " + str(govname[i]) +"  ,  ")
                    Cname = analysis(FraudulentName,mydata,family,Suspect,Familynames,"PERSON")
                    if Cname == []:
                        Cname = analysis(cleannames(Orgname),mydata,family,Suspect,Familynames,"ORGANIZATION")
                    Cname = list(set(Cname) - set(fulname(Cname,re.findall('(%s), \d+, %s'%('|'.join(lastname(Cname)),'was killed'),mydata))))
                    Remove = re.findall('(%s) (%s)'%('|'.join(lastname(Cname)),'died|dead|suffered'),mydata)                    
                    if Remove != [] and Remove[0][1] != "suffered" and fulname(Cname,[Remove[0][0]])[0] not in Suspect:
                        Cname = list(set(Cname) - set(fulname(Cname,[Remove[0][0]])))
                    if Aliasname != [] and re.findall('|'.join(Realname),'|'.join(Cname)) !=[]:
                        for i in range(len(set(re.findall('|'.join(Realname),'|'.join(Cname))))):
                            RealName = ''.join(Realname[i])
                            AlaisName = ' , '.join(Aliasname[i])
                            Cname = re.sub('%s'%RealName,'%s( %s )'%(RealName,AlaisName),'//'.join(Cname)).split('//')
                    FraudlentName = "Fraudulent Name :" + "\t" + str(Cname)+"\n"
                    print "Fraudulent Name :" + "\t" + str(Cname)+"\n"
                    if State != []:        
                        for i in range(len(Cname)):
                            LocName.append("State: " + State +" , " + "City:" + City+"  :  ")
                        print "State: " + State +" , " + "City:" + City
                    else:
                        print "Location not found"
                    scheme = schemee(words)
                    SchemeName= "Category:" +"\t" + str(scheme)
                    print SchemeName
                    writedata(mydat,FraudlentName,LocName,scheme,govname,GovName,Cname)
                    #EXCEL OUTPUT
                    s.write(e,1,FraudlentName)
                    s.write(e,2,LocName)         
                    s.write(e,3,GovName)
                    s.write(e,4,SchemeName)
                    wb.save('Verisk iso.xls')
                except:
                    pass
            else:
                print "No Criminal activity"
        elif criminalactivity == "True" and Databaseflag == "True":
            print Fraudlentname+"\n\n"+StateandCity+"\n\n"+ Govofficials+"\n\n"+Scheme
            s.write(e,1,Fraudlentname)
            s.write(e,2,StateandCity)         
            s.write(e,3,Govofficials)
            s.write(e,4,Scheme)
            wb.save('Verisk iso.xls')
        else:
            print "No crime occured"
            s.write(e,1,'No Crime Occured')
            wb.save('Verisk iso.xls')

#User interface to interact with the user to input the article number for analysis.
def getinput(len):
    print "Would you like to run Single or Multiple article?(S/M)"
    num = raw_input()
    if num in ["s","S"]:        
        print "Enter the cell number that you would like to run"
        try:
            sn1=int(raw_input()) -1
            sn2 = sn1 + 1
        except ValueError:
            print("That's not an integer input!")
    elif num in ["m","M"]:        
        print "If you would like to run the analysis for all article type Y else N"
        all = raw_input()  
        if all in ["Y","y"]:
            sn1 = 1
            sn2 = len
        elif all in ["N","n"]:
            print "Enter the starting cell number"        
            try:
                sn1=int(raw_input()) -1
            except ValueError:
                print("That's not an integer input!")
            print "Enter the ending cell number"
            try:
                sn2=int(raw_input())
            except ValueError:
                print("That's not an integer input!")
        else:
            print "Enter 'Y' for all or 'N' for user choice"
    else:
        print "Enter 'S' for Single or 'M' for Multiple"
    return sn1,sn2    

#Clean the data
def cleanrawdata(mydata):
   #Uses only alphabets, numbers and special character
    mydata = re.sub('\v+','. ',mydata)
    a = ''.join([i if ord(i) < 127 and ord(i) > 31 else '' for i in mydata])
    y=re.findall(r"[\w]+|[.,!?;':]",a)
    a=''
    for i in range(len(y)):
        if y[i].isalnum():
            a = a +" "+y[i]
        else:
            a = a + y[i]
    a = re.sub('\.+','.',a)
    return a

#Identify all the names in an article using the Stanford NER tool kit.
def name(dictionary,mydata,Familyflag):
    tagger = ner.SocketNER(host='localhost', port=8080)
    Noun = tagger.get_entities(mydata)
    Fullname =initialze() 
    Allsent = '\n'.join(tokenize.sent_tokenize('\n'+mydata+'\n'))
    a=[];c=[];Realname=[];Aliasname=[];Temp=[];Suspect = []
    try:
        Pnames = list(set(Noun["PERSON"]))
        if re.findall(Aname,mydata) !=[]:
            f = re.findall(Aname,mydata)[0]
            doc=re.findall('\\n(.*?) (%s .*?)\\n'%f,Allsent)
            for e in doc:
                Username = e[0];Fakename = e[1] 
                Realname += fulname(Pnames,[re.findall('|'.join(lastname(cleannames(Pnames))),Username)[-1]])
                Summa = [i for i in tagger.get_entities(Fakename)["PERSON"] if i in Pnames and len(i.split()) >1]
                Aliasname.append(Summa)
                Temp += Summa
        Pnames = list(set(Pnames) - set(Temp))
        Pname,Suspect = Suspectednames(Pnames,mydata,tagger)
        Suspect += cleannames(Realname)       
        d=list(set(' '.join(Suspect+Temp).split()))
        for i in Pname:
            for j in i.split():
                if j not in d and len(i.split()) > 1 and mydata.split()[1].title() != i.split()[1]:
                    d.append(j)
                    a.append(i)
            if len(i.split()) == 1 and re.findall(i,'|'.join(d)) == [] and mydata.split()[1].title() != i:
                if Familyflag == "True":
                    fpattern = re.findall('\\n(.*?%s,? %s.*?)\\n'%(vFamily,i),Allsent)
                    if fpattern !=[]:
                        famname = tagger.get_entities(''.join(fpattern))["PERSON"][0].split()[-1]
                        if famname != i:
                            c.append(i+' '+famname)
                            familydict[i] = (i+' '+famname)
                        else:
                            c.append(i +' ' +re.findall('%s'%'|'.join(lastname(a)),fpattern[0])[-1])
                            familydict[i] = (i+' '+re.findall('%s'%'|'.join(lastname(a)),fpattern[0])[-1])
                    ffname = re.findall('(%s) and (%s) '%(i,'|'.join(Pname)),mydata)
                    if ffname != []:
                        if ffname[0][0] not in d and (ffname[0][1].split()) > 1 and len(ffname[0][-1].split()) > 1:
                            c.append(ffname[0][0]+' '+ffname[0][1].split()[-1])
                            familydict[ffname[0][0]] = ffname[0][0]+' '+ffname[0][1].split()[-1]
                    findfname = re.findall('\\n(.*?)(%s) (%s).*?\\n'%(vFamily,i),Allsent)
                    if findfname != []:
                        familylname = re.findall('%s'%'|'.join(lastname(a)),findfname[0][0])
                        if familylname !=[]:
                            c.append(i+' '+familylname[0])
                            familydict[i] = i+' '+familylname[0]
        Fullname = list(set(a))
        Fullname = cleannames(Fullname)
        #Alias names
        #re.findall('[%s],? (%s),? and?(\w+?)'%(Aname,'|'.join(Name)),mydata)
    except:
        Fullname = ["No Person Name available"]
    try:
        Name = list(set(Noun["ORGANIZATION"]))
        # Using ner extract the Name, Location and the Organization details from the text.
        x =initialze()
        for i in Name:
            org = re.findall(r'\w+', i)
            if len(org)>1 and len(org) < 4 and dictionary[org[-1]] >3:
                x.append(i)
    except:
        x= ["No Org Name available"]
    Suspect = list(set(Suspect) - set(re.findall('%s (%s)'%(unwantednames,'|'.join(Suspect)),mydata)))
    unwanted = re.findall('%s (%s)'%(unwantednames,'|'.join(Fullname)),mydata)
    Fullname = list(set(Fullname)-set(re.findall('|'.join(Fullname),'|'.join(x))+unwanted))
    return (Fullname,x,Noun,Suspect,c,cleannames(Realname),Aliasname)

#clearing the junk values and mis-spelled names 
def cleannames(Pname):
    Names=initialze()
    for i in Pname:
        shape = '(?:^|(?<= ))(%s)(?:(?= )|$)'
        if re.findall(shape%vhatecrime,i) == [] and (i.istitle() or list(i.split()[-1])[1].islower())  and \
            len(re.findall(r"[\w]",i)) > 5 and re.findall(shape%wrongnamepattern,i) == [] and \
            len(i.split()) < 5 and re.findall("\d", i) == []:
                Names.append(i)
    Names.sort(key=len, reverse=True)
    return Names

#identifying the Suspects from the list
def Suspectednames(Pnames,mydata,tagger):
    Pname=initialze();Suspect = initialze()
    for i in Pnames:
        shape = '(?:^|(?<= ))(%s)(?:(?= )|$)'
        if re.findall(shape%vhatecrime,i) == []  and len(i.split()) < 5 and re.findall("\d", i) == [] \
        and len(re.findall(r"[\w]",i)) > 4 and re.findall(shape%wrongnamepattern,i) == [] :
            Pname.append(i)
    Pname.sort(key=len, reverse=True)
    for i in Pname:
        if len(i.split()) > 3 and len(i.split()[0]) <3 and len(i.split()[1]) <3:
            Pname.append(' '.join(i.split()[2:len(i.split())]))
            Pname.remove(i)
        if (''.join(re.findall("(%s),? \d+,.*?[%s]"%(i,vFinal),mydata)) in Pname or 
        ((re.findall("(%s) (%s)"%(i,Negwords),mydata) != [] or re.findall("(%s) (%s)"%(Negwords,i),mydata) != []) and \
        re.findall("(%s) (%s)"%(vhatecrime,i),mydata) == [])or ''.join(re.findall("(%s), age \d+,.*?[%s]"%(i,vFinal),mydata)) in Pname):
            if len(i.split())>1:            
                Suspect.append(i)
                continue
            elif len(fulname(Pname,[i])[0].split()) >1:
                Suspect += fulname(Pname,[i])
                continue
        if (re.findall('(%s),.*?(%s)'%(i,Negwords),'\n'.join(tokenize.sent_tokenize(mydata))) != [] and \
        re.findall("(%s) (%s)"%(vhatecrime,i),mydata) == [])and len(i.split())>1:
            findword=re.findall('(%s),.*?(%s)'%(i,Negwords),mydata)[0]
            forward = re.findall('%s,(.*?)%s'%(findword[0],findword[1]),mydata)[0]
            if re.findall('|'.join(' '.join(Pname).split())+"|"+vhatecrime,forward) == [] and re.findall('%s'%vFamily,forward) == []:
                Suspect.append(i)
                continue
        FLname = re.findall("(%s,? \d+,).*? (%s)"%(i,Negwords),mydata)
        if FLname != [] and len(i.split())  == 1:
            Fname = cleannames(tagger.get_entities(mydata.split(FLname[0][0])[0])["PERSON"])
            Suspect += [F for F in Fname if lastname([F]) == i]
            continue
        if (re.findall("(%s),? (%s)"%(i,Negwords),mydata) != [] or re.findall("(%s) was (%s)"%(i,Negwords),mydata) != []) and \
        len(i.split())  == 1 and re.findall("(%s) (%s)"%(vhatecrime,i),mydata) == []:
            Suspect += cleannames(fulname(Pname,[i]))
            continue
        if ''.join(re.findall("(%s), \d+, of"%(i),mydata)) in Pname:
            Suspect.append(i)
            continue
        try:
            if len(i.split())>1:
                Suspect.append(re.findall(i,re.sub(r'\n','|',open("Fraudlentname.txt", "r").read()))[0])
                open("Fraudlentname.txt", "r").close()
        except:
            pass
    maybe = re.findall("%s, (%s)(.*?) and (%s)"%(multicriminal,'|'.join(Pname),'|'.join(Pname)),mydata)
    if maybe != []:
        Suspect += [i for i in re.sub(', ','|','|'.join(maybe[0])).split('|') if i in Pname]
    Pname = list(set(Pname)-set(Suspect))
    Pname.sort(key=len, reverse=True)
    return (Pname, Suspect)
#Final check on criminal activities involved in the article.
def Fcheck(mydata):
    cfr = "False"
    family = "False"
    #Using dictionary to identify crime occurrence in the article.
    if re.findall(vFinal,mydata) !=[]:
        cfr ="True"
    #Family flag as true to indicate there will be multiple names involved in the article with the same last name.
    if re.findall(vFamily,mydata) !=[]:
        family = "True"
    return cfr, family

#Identify the name(s) of the government officals and their corresponding title
def findname(Fullname,mydata,crimetitle,words,Lname):
    ctitle=initialze()
    gname =initialze()
    for i in Fullname:
        data = mydata
        for j in crimetitle:
            Flag = 0
            try:
                #Identifies the Name of the person in government position or proved as innocent in the crime.
                #sentence following the name till the crimetitle
                forward = mydata.split(i)[1].split(j)[0]
                #sentence prior to the name till the crimetitle
                backward = mydata.split(i)[0].split(j)[-1]
                if backward.split() != []:
                    #when the length of backward and forward is equal priority is provided to the forward if it's immediately
                    #followed by a comma or when the backward reaches the first letter of the article.
                    if (mydata.split()[0] == backward.split()[0] or (len(forward) == len(backward) and forward[0] == ",")) :
                        backward = mydata
                    try:
                    #When the forward reaches the last letter priority is given to the backward text.
                        if mydata.split()[-1] == forward.split()[-1]:
                            forward = mydata
                    except:
                        pass
                else:
                    if mydata.split()[0] == i.split()[0]:
                        backward = mydata
                #Using sentimental analysis to identify the negativity of the text
                if len(forward) < len(data) and len(forward) < len(backward) and senty(forward)[1] ==0 \
                and re.findall('%s'%wrongtitlepattern,forward) == []:
                    words = re.findall(r"[\w']+|[.,!?; ]", forward) 
                    #Eliminating the possibility of criminal name getting added to a government title.
                    for word in words:
                        if word in Lname:
                            Flag = 1
                    if Flag == 0:
                        data = forward
                        n = j
                #Using sentimental analysis to identify the negativity of the text
                if len(backward) < len(data) and len(forward) > len(backward) and senty(backward)[1] ==0 \
                and re.findall('%s'%wrongtitlepattern,backward) == []:
                    words = re.findall(r"[\w']+|[.,!?; ]", backward) 
                    #Eliminating the possibility of criminal name getting added to a government title.
                    for word in words:
                        if word in Lname:
                            Flag = 1
                    if Flag == 0:
                        data = backward
                        n = j
            except:
                pass
        try:
            if len(mydata) != len(data):
                gname.append(i)
                ctitle.append(n)
                if len(forward)<len(backward):
                    mydata = ''.join(re.split(i+data+n,mydata))
                else:
                    mydata = ''.join(re.split(n+data+i,mydata))
        except:
            continue
    return (ctitle, gname,mydata)

#Identify the Full name of a person using the Last name
def fulname(Name,Lname):
    z=initialze()
    for j in Lname:
        n=j
        for i in Name:
            try:
                if re.findall(r"[\w]+",i)[-1] == j and len(re.findall(r"[\w]",i)) > len(re.findall(r"[\w]",n)):                 
                    n=i
                elif re.findall(r"[\w]+",i)[-2] == j and len(re.findall(r"[\w]",i)) > len(re.findall(r"[\w]",n)):
                    n=i
            except:
                pass
        z.append(n)
    return list(set(z))

#Identify the location of the fraud
def location(Noun,mydata):
    try:
        Location = list(set(Noun["LOCATION"]))
        Location.append(mydata.split( )[0])
        #Using google to identify the location
        g = geocoder.google(','.join(Location))
        if g.city == None:
            g.city ="Not known"
        return g.state, g.city
        
    except:
        return ([],[])

#Initializing value
def initialze():
   z=[]
   return z

#To identify the last name of a person
def lastname(Name):
    y =initialze()
    for i in Name:
        if re.findall(r'\w+',i)[-1] in suffix:
            y.append(re.findall(r"[\w]+",i)[-2])
        else:
            y.append(re.findall(r"[\w]+",i)[-1])
    return list(set(y))
	
#Identify the name of the fraudulent involved.
def analysis(Name,mydata,Familyflag,Suspect,Familynames,shift):
    tagger = ner.SocketNER(host='localhost', port=8080)
    sentences = tokenize.sent_tokenize(mydata)
    names=initialze();pronoun =initialze()
    Lname = lastname(Name)
    #Tagging out the sentences with the non criminal activities.
    for i in sentences:
        #shape = '(?:^|(?<= ))(%s)(?:(?= )|$)'
        if re.findall('.*?(%s).*?'%'|'.join(ihatecrime),i,re.IGNORECASE) == [] and re.findall('.*?(%s).*?'%action,i) == [] \
        and senty(i)[1] > 0 and re.findall("%s"%('|'.join(Lname)),i,re.IGNORECASE) != []:
            try:  
                pronoun = tagger.get_entities(i)[shift]
                for pnoun in pronoun:
                    if (i == sentences[0] and (re.findall('[By|by] (%s)'%pnoun,i) != [] or \
                    re.findall('CONTACT: (%s)'%pnoun,i) != [])) or re.findall('(%s).*?(%s)'%(pnoun,options),i) != []:
                            Name = list(set(Name) -set([pnoun])) 
                    elif pnoun in Name or pnoun.title() in Lname:
                        names.append(pnoun.title())  
                    elif pnoun in Name or pnoun in Lname:
                        names.append(pnoun)
                    elif len(pnoun.split()) == 1 and Familyflag == "True" and \
                    pnoun in ' '.join(Familynames).split():
                        names.append(familydict[pnoun])
                                                
            except:
                pass
    Lastname = list(set(names+Suspect))   
    Fullname = fulname(Name+Suspect,Lastname)  
    Fullname.sort(key=len, reverse=True)
    a=[];d=[]
    for i in Fullname:
        for j in i.split():
            if j not in d and len(i.split()) > 1:
                d.append(j)
                a.append(i)
    Fullname = list(set(a))
    Famname = initialze()
    #Identifying the family member involved in the crime
    if Familyflag == "True":
        for i in Fullname:
            if i.split()[-1] in Lastname and i not in Fullname:
                forward = mydata.split(i)[1].split(". ")[0]
                backward = mydata.split(i)[0].split(". ")[-1]            
                if senty(forward)[1] != 0 or senty(backward)[1] != 0:
                    Famname.append(i)
        Fullname = Fullname+Famname   
    return Fullname        

#Function to identify the type of scam in an article.
def schemee(words):
    scheme=initialze()   
    w=[item.lower() for item in words]
    w=list(set(w))
    for word in w:
        if word in All:
            if word in Medicare:
                scheme.append("Medicare")
            if word in Medicaid:
                scheme.append("Medicaid")
            if word in Health:
                scheme.append("Health")
            if word in Life:
                scheme.append("Life")
            if word in Liability:
                scheme.append("Liability")
            if word in Disability:
                scheme .append("Disability")
            if word in Worker:
                for item in w:
                    if item in Injured:
                        scheme.append("Worker Compensation")
            if word in Payroll:
                scheme.append("Business")
            if word in Auto:
                scheme.append("Auto") 
            if word in Homeowners:
                scheme.append("Homeowners")
    if scheme == []:
        scheme = ["Others"]
    if "Medicare" in scheme and "Health" in scheme:
        scheme = ["Medicare"]
    if "Medicaid" in scheme and "Health" in scheme:
        scheme = ["Medicaid"]
    if "Disability" in scheme and "Worker Compensation" in scheme:
        scheme = ["Worker Compensation"]
    return list(set(scheme))

#Creating a Database
def databasefile(mydata):
    Fraudlentname = "NULL"
    StateandCity = "NULL"
    Scheme = "NULL"
    Govofficials = "NULL"
    Databaseflag = "False"
    file_location = "C:\Users\Vivek\Documents\Python Scripts\databasefile.xls" #change the file location
    try:
        database_worksheet = xlrd.open_workbook(file_location)
        workbookdata_sheet = database_worksheet.sheet_by_index(0)
        for r in range(workbookdata_sheet.nrows):
            try:
                if workbookdata_sheet.cell_value(r,0) == mydata:
                    Fraudlentname = workbookdata_sheet.cell_value(r,1)
                    StateandCity = workbookdata_sheet.cell_value(r,2)
                    Scheme = workbookdata_sheet.cell_value(r,3)
                    Govofficials = workbookdata_sheet.cell_value(r,4)
                    Databaseflag = "True"
                    break
            except:
                pass
    except:
        worksheet = xlwt.Workbook()
        database_worksheet = worksheet.add_sheet('Summary')
        database_worksheet.write(0,0,"Case Note")
        database_worksheet.write(0,1,"Fraudulent Name")
        database_worksheet.write(0,2,"State & City")
        database_worksheet.write(0,3,"Scheme")
        database_worksheet.write(0,4,"Government Officials")
        worksheet.save('databasefile.xls')
    return Fraudlentname,StateandCity,Scheme,Govofficials,Databaseflag

# Recording the details for future analysis  
def writedata(mydata,Fraudlentname,StateandCity,Scheme,Govofficials,GovName,Cname):
    file_location = "C:\Users\Vivek\Documents\Python Scripts\databasefile.xls"
    workbookdata = xlrd.open_workbook(file_location,formatting_info=True)
    wbook = copy(workbookdata)
    database_worksheet = wbook.get_sheet(0)
    workbookdata_sheet = workbookdata.sheet_by_index(0)
    r = workbookdata_sheet.nrows
    database_worksheet.write(r,0,mydata)
    database_worksheet.write(r,1,Fraudlentname)
    database_worksheet.write(r,2,StateandCity)
    database_worksheet.write(r,3,Scheme)
    database_worksheet.write(r,4,GovName)
    wbook.save('databasefile.xls')
    with open("Fraudlentname.txt", "a") as text_file:
        for Names in Cname:
            text_file.write("%s\n" %Names)
    text_file.close()
    with open("Govofficials.txt", "a") as text_file:
        for Names in Govofficials:
            text_file.write("%s\n" %Names)
    text_file.close()

#Function to identify multiple title users
def plural(mydata,Name):
    title = [];names=[];Names = Name
    crimeofficerlist = list(set(re.findall(pluralhatecrime,mydata,re.IGNORECASE)))  
    for sig in crimeofficerlist:
        a = re.findall("%s (%s)(.*?), and (%s)"%(sig,'|'.join(Name),'|'.join(Name)),mydata)
        if a!=[]:
            names+=[i for i in re.sub(', ','|','|'.join(a[0])).split('|') if i in Name]
            title.append(sig)
            Names = list(set(Name)-set(names))
            mydata = ''.join(re.split(re.findall('(%s.*?%s)'%(sig,names[-1]),mydata)[0],mydata))
    return title,names,Names,mydata


#List to identify the type of scam 
Medicare = ["medicare"]
Medicaid = ["medicaid"]
Life = ["life"]
Health = ["health"]
Liability = ["liability"]
Disability = ["disability"]
Homeowners = ["home"]
Worker = ["worker","workers","employee","labour","labourer","work"]
Injured = ["injured","injury","disease","impairment","infirmity","compensation","comp"]
Payroll = ["payroll","wage","payment","overbilling","loan"]
Auto = ["auto","driver","vehicle","motor","passenger","wheels"]
All = Medicare+Medicaid+Life+Liability+Disability+Worker+Injured+Payroll+Auto+Health

#List of government offical designation
ihatecrime = ["Commissioner","attorney","prosecutor","Honorable","Officer","Judge","Prosecutor","Attorney",
"Mayor","Director","director","Investigator","spokesman","investigator","professor","spokeswoman","Sheriff",
"lawyer","auditor","Superintendent","regulator","whistleblower","Atty.","secretary for","cosmetic giant",
"Detective","appraiser","firefighter","marshal","dismissed","senior special agent","Emergency Management",
"Army Veteran","Reporter","federal prosecutor","special agent","District Justice","Justice",
"corporate special investigations manager","writer","editor","District Justice","Magistrate","economist",
"workers' comp fraud bureau","investigations manager","special agent","Senator","District Attorney",
"Superintendent","investigator","detective","Secretary of State","gubernatorial","general counsel",
"Insurance Fraud Bureau","his lawsuit","Insurance Fraud Prevention Unit","policeman-","fire chief"
]
#Secondary list
capcrime = r"CEO|FIP|FBI|Sgt|Gov\.|Lt\.|NAII|Supt\.|Sen\."
#identify multiple criminal names
multicriminal = "defendants"
vhatecrime = '|'.join(ihatecrime)
#To identify multiple government officials
pluralhatecrime = "Investigators|writers"
#Dictionary of the words used to identify the crime rate of an article
Final = ["pleaded","charged","sentenced","convicted","arrested","accused","victimized","sued","uninsured","to break the law",
"indicted","pleading","plead","indict","arrest","alerted","unlicensed","victimized","duped","surrendered","fled"
"indictment","alleged","cheating","pled","trampled","apprehended","conspired","torched","acquitted","revoked"]
vFinal = '|'.join(Final)
#Identify important criminal names
Negwords = "pleaded|charged|sentenced|convicted|arrested|accused|indicted|surrendered|apprehended|torched|acquitted|fled \
|to break the law|paid back the money as part of a settlement|diverted payroll deductions "
wrongtitlepattern = "appeared before"
#To eliminate unwanted naming patterns
unwantednames = "By|, of"
#Dictionary to identify family members involved in the crime.
Family =["father","brother","sister","wife","husband","mother","son","daughter","aunt","nephew","niece","uncle",
"brothers","sisters","sons","daughters"]
vFamily = '|'.join(Family)
#dead person names
options = 'died|wanted relief'
#Eliminate witness
verb = ["said","says","CONTACT:","innocent","released"]
action = '|'.join(verb)
#Eliminate Juniour or Senior from the last name during Analysis
suffix =["Jr","Sr","jr","sr"]
#To elimiinate wrong naming patterns
wrongnamepattern = "County|Street|Company|Road|Office|Limited|St.|Ave.|Avenue|U. S.|Attorney|Staff|Friday|Union|Navigator|\
Would"
#To identify alternate nammes
Aname = "aka|alias|better known by|AKA|also known as|also called|otherwise known as|commonly known as"
#Creating a family dictionary
familydict={}
#Start
main()
