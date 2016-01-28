# VERISK competition to identify the fradulant name, location and scheme
# Vivek Pandiyan
# Project start date: Dec 11th December, 2015

#Library imported for identifying the pronouns
from nltk.tag import pos_tag
#Library imported for sentimental analysis
from vaderSentiment.vaderSentiment import sentiment as senty
#Library used for importing counter for identifying the frequency
from collections import Counter
#Library imported to input excel file
import xlrd
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
    tagger = ner.SocketNER(host='localhost', port=8080)
    file_location = "C:/Users/Vivek/Desktop/Class Notes/Python/Project_PY/ISO/ISO.xlsx"
    workbook = xlrd.open_workbook(file_location)
    print "Enter the sheet number"
    try:
        w=int(raw_input()) -1
    except ValueError:
        print("That's not an integer input!")
    sheet = workbook.sheet_by_index(w)
    z=initialze()
    data = [[sheet.cell_value(r,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    n,m=getinput(len(data))
    wb = copy(workbook)
    s = wb.get_sheet(w)
    for e in range(len(data))[n:m]:
        print "-"*20 
        print "CASE NUMBER %r" %(e)
        print "-"*20 
        GovName = initialze()
        LocName = initialze()
        p = initialze()
        mydat = data[e][0]
        mydata = cleanrawdata(mydat)
        words = re.findall(r'\w+', mydata)
        criminalactivity,family= Fcheck(words)
        if criminalactivity == "True":
            #dict = count(words)
            word_counts = Counter(words) # Count the words
            Noun = tagger.get_entities(mydata)
            Name,Orgname = name(Noun,word_counts)
            State, City = location(Noun,mydata)
            z=initialze()
            for i in words:
                if i in ihatecrime:
                    z.append(i)		
            crimetitle=set(z)
            if State == [] and City == [] and Name == "No Person Name available":
                print "No Criminal activity"
            try:
                Lname = lastname(Name)
                Fullname = fulname(Name,Lname)
                govtitle, govname = findname(Fullname,mydata,crimetitle,words,Lname)
                splitname= initialze()
                for i in govname:
                    for j in re.findall(r"[\w']+",i):
                        splitname.append(j)
                if govtitle != []:
                    for i in range(len(govtitle)):
                        print govtitle[i] + ":" + govname[i] +"\n" 
                        GovName.append(govtitle[i] + " : " + govname[i] +"  ,  ")
                FraudulentName = list(set(Name) -set(govname))
                sentences = mydata.split(". ")
                vsi,para = senti(sentences)
                Cname = analysis(vsi,para,splitname,FraudulentName,mydata,family)
                if Cname == []:
                    Cname = analysis(vsi,para,splitname,Orgname,mydata,family)
                FraudlentName = "Fraudulent Name :" + "\t" + str(Cname)+"\n"
                print "Fraudulent Name :" + "\t" + str(Cname)+"\n"
                if State != []:        
                    for i in range(len(Fullname)):
                        LocName.append("State: " + State +" , " + "City:" + City+"  ,  ")
                    print "State: " + State +" , " + "City:" + City
                else:
                    print "Location not found"
                scheme = schemee(words)
                SchemeName= "Category:" +"\t" + str(scheme)
                print "Category:" +"\t" + str(scheme)
                #TEXT OUTPUT                
                with open("Output.txt", "a") as text_file:
                    text_file.write("-"*20)
                    text_file.write("\nCASE NUMBER: %r\n" % e)
                    text_file.write("-"*20)
                    text_file.write("\nState: %s City: %s\n Criminal Name: %s\n Scheme: %s\n" % (State,City,Fullname,scheme))
                #EXCEL OUTPUT
                s.write(e,1,FraudlentName)
                s.write(e,2,LocName)         
                s.write(e,3,GovName)
                s.write(e,4,SchemeName)
                wb.save('Verisk iso.xls')
            except:
                pass
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
	a = ''.join([i if ord(i) < 127 and ord(i) > 31 else '' for i in mydata])
	y=re.findall(r"[\w']+|[.,!?;]",a)
	a=''
	for i in range(len(y)):
		if y[i].isalnum():
				a = a +" "+y[i]
		else:
				a = a + y[i]
	return a

#Identify all the names in an article using the Stanford NER tool kit.
def name(Noun,dict):
    try:
        Names=initialze()
        Pname = list(set(Noun["PERSON"]))
        for i in Pname:
            Flag=0
            for h in i.split():
                if h in ihatecrime or h.islower() or len(re.findall(r"[\w]",i)) < 4:
                    Flag = 1
            if Flag == 0:
                a=' '.join(j for j in i.split() if not j.isdigit())
                Names.append(a)
    except:
        Names = "No Person Name available"
    try:
        Name = list(set(Noun["ORGANIZATION"]))
        # Using ner extract the Name, Location and the Organization details from the text.
        x =initialze()
        for i in Name:
            org = re.findall(r'\w+', i)
            if len(org)>1 and len(org) < 4 and dict[org[-1]] >3:
                x.append(i)

    except:
        x= ["No Org Name available"]
    return Names,x

#Final check on criminal activities involved in the article.
def Fcheck(words):
    cfr = "False"
    family = "False"
    for i in words:
        #Using dictionary to identify crime occurrence in the article.
        if i in Final:
            cfr ="True"
        #Family flag as true to indicate there will be multiple names involved in the article with the same last name.
        if i in Family:
            family = "True"
    return cfr, family

#Identify the name(s) of the government officals and their corresponding title
def findname(Fullname,mydata,crimetitle,words,Lname):
    ctitle=initialze()
    gname =initialze()
    for i in Fullname:
        data = mydata
        Flag = 0
        for j in crimetitle:
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
            if len(forward) < len(data) and len(forward) < len(backward) and senty(forward)[1] ==0:
                words = re.findall(r"[\w']+|[.,!?; ]", forward) 
                #Eliminating the possibility of criminal name getting added to a government title.
                for word in words:
                    if word in Lname:
                        Flag = 1
                if Flag == 0:
                    data = forward
                    n = j
            #Using sentimental analysis to identify the negativity of the text
            elif len(backward) < len(data) and len(forward) > len(backward) and senty(backward)[1] ==0:
                words = re.findall(r"[\w']+|[.,!?; ]", backward) 
                #Eliminating the possibility of criminal name getting added to a government title.
                for word in words:
                    if word in Lname:
                        Flag = 1
                if Flag == 0:
                    data = backward
                    n = j
        try:
            if len(mydata) != len(data):
                gname.append(i)
                ctitle.append(n)
        except:
            continue
    return (ctitle, gname)

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
        Location.append(mydata.split(' ', 1)[0])
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
	
#Identify the negativity of a sentence
def senti(sentences):
	vsi=[]
	para =[]
	for sentence in sentences:
		if senty(sentence)[1] > 0.0:
			para.append(sentence)
			vsi.append(senty(sentence)[1])
	return vsi,para

#Identify the name of the fraudulent involved.
def analysis(vsi,para,splitname,Name,mydata,Familyflag):
    ind = sorted(vsi,reverse=True)
    F=initialze()
    names=initialze()
    Lname = lastname(Name)
    #Tagging out the sentences with the non criminal activities.
    for i in ind:
        speech = re.findall(r'\w+', para[vsi.index(i)])
        for j in speech:
            if j in ihatecrime or j in verb:
                Flag = 0
                break
            else:
                Flag = 1
        z = [Flag] + [vsi.index(i)]
        F.append(z)
    for i in F:
        if i[0] == 1:
            #Using NLTK to identify the names by crossing with the NER data to identify the fraudlent name.
            z=pos_tag(re.findall(r"[\w]+",para[i[1]]))
            pronoun = [word for word,pos in z if pos == 'NNP' or pos == 'RB']
            for pnoun in pronoun:
                if pnoun in Lname:
                    if pnoun not in splitname:
                        names.append(pnoun)
    Lastname = list(set(names))   
    Fullname = fulname(Name,Lastname)
    Famname = initialze()
    #Identifying the family member involved in the crime
    if Familyflag == "True":
        for i in Name:
            if i.split()[-1] in Lastname and i not in Fullname:
                forward = mydata.split(i)[1].split(". ")[0]
                backward = mydata.split(i)[0].split(". ")[-1]            
                if senty(forward) != 0 or senty(backward) != 0:
                    Famname.append(i)
        Fullname = Fullname+Famname   
    # Eliminate the redundant Name occurrences
    Fullname.sort(key=len, reverse=True)
    a=[]
    b=[]
    c=[]
    d=[]
    for i in Fullname:
        if len(i.split()) > 2:
            for j in i.split():
                d.append(j)
                c.append(i)
        elif len(i.split()) == 2:
            for j in i.split():
                if j not in d:
                    d.append(j)
                    a.append(i)
        elif len(i.split()) == 1:
            if i not in d:
                b.append(i)        
    Fill = list(set(c)) + list(set(b)) + list(set(a))
    return Fill         

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
All = Medicare + Medicaid+Life+Liability+Disability+Worker+Injured+Payroll+Auto+Health

#List of government offical designation
ihatecrime = ["Commissioner","attorney","prosecutor","Honorable","Officer","Judge","Prosecutor","Attorney",
"Mayor","Director","director","Investigator","spokesman","investigator","professor","spokeswoman","agent","Sheriff",
"lawyer","auditor","CEO","Superintendent","Chief","regulator","innocent","FIP","FBI","Sgt",
"Detective","Lt","NAII","appraiser","firefighter","marshal","dismissed"]

#Dictionary of the words used to identify the crime rate of an article
Final = ["pleaded","charged","sentenced","convicted","arrested","accused","victimized","sued","uninsured",
"indicted","pleading","plead","indict","arrest","alerted","unlicensed","victimized","duped","surrendered",
"indictment","alleged","cheating","pled","trampled"]

#Dictionary to identify family members involved in the crime.
Family =["father","brother","sister","wife","mother","son","daughter","brothers","sisters"]

#Eliminate witness
verb = ["said"]

#Eliminate Juniour or Senior from the last name during Analysis
suffix =["Jr","Sr","jr","sr"]

#Start
main()