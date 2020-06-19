#Aceasta varianta este dual branded


import zipfile, os.path, time, shutil, smtplib, csv, signature
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import timedelta, date, datetime
import MySQLdb


DB_connected = False
try:
    db = MySQLdb.connect(host="127.0.0.1",
                     user="root",
                     passwd="",
                     db="audit")
    cur = db.cursor()
    DB_connected = True
except (Exception) as error:
    print (error)


def daterange(date1, date2):
    for n in range(int ((date2 - date1).days)+1):
        yield date1 + timedelta(n)

        
def resetare_variabile(brand):
    CONTRACTE.cel_putin_1_procesare = False
    global mesaj, Lista, EOF, ListaRapoarteNumeric, Lista_temp, Lista_obiecte, StartPoint, EndPoint, VerificareInteger, ListaRapoarte
    mesaj=''
    Lista=[]
    EOF = "                                      *                    END OF REPORT                   *                                       "

    ListaRapoarteNumeric = [[], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], []]
    Lista_temp = []
    Lista_obiecte = []
    StartPoint = []
    EndPoint = []
    VerificareInteger = 0
    ListaRapoarte = ["CASH RENTALS WITH A CHARGE MODIFIED AFTER CHECKIN",
                     "RENTALS WHERE TIME AND DISTANCE HAS BEEN ADJUSTED",
                     "CHECKIN DELAY ANALYSIS",
                     "RENTALS WHERE INSURANCES HAVE BEEN MODIFIED",
                     "CASH RENTALS WHERE RATE CODE MODIFIED AFTER CHECKOUT",
                     "LOW VALUE MANUAL RATE RENTALS ",
                     "LOW VALUE MANUAL RATE RENTALS ",
                     "LOW VALUE MANUAL RATE RENTALS ",
                     "NO CHARGE RENTALS",
                     "UNAUTHORISED DIRECT BILLING RENTALS",
                     "EMPLOYEE RENTALS",
                     "RENTALS WITH NO FUEL CHARGE",
                     "LONG DURATION RENTALS",
                     "LONG DISTANCE RENTALS",
                     "CHECKIN DELAY STATISTICS",
                     "RENTALS WHERE COMMISSION HAS BEEN MODIFIED",
                     "MULTIPLE USE OF FTP NUMBERS"]


months={"1":"JANUARY   ","2":"FEBRUARY  ","3":"MARCH     ","4":"APRIL     ","5":"MAY       ","6":"JUNE      ","7":"JULY      ","8":"AUGUST    ","9":"SEPTEMBER ","10":"OCTOBER   ","11":"NOVEMBER  ","0":"DECEMBER  "}


def scriere_log(mesaj):
    LogFile=open("LogFileAudit.txt","a")
    LogFile.write(mesaj)
    print (mesaj)
    LogFile.close()

def citire_fisier():
    global Lista, rawText
    for e in rawText:  # Citire fisier de interpretat si inserare fiecare rand in lista, scazand in acelasi timp primul caracter din fiecare rand
        j = 1
        rand = ""
        while j < len(e):
            rand = rand + e[j]
            j += 1
        Lista.append(rand)
    rawText.close()    

def main1(Brand):
    global Fisier, rawText
    if Brand=="brand1":
        cale="EPROLD\\XPSBX\\FTP\\SRAA02P\\CR"
    else:
        cale="EEROLD\\XPSBX\\FTP\\SRAA02P\\CR"
    if os.path.isfile(Fisier):
        try:
            zip=zipfile.ZipFile(Fisier)
            zip.extractall()
            rawText=open(cale,"r")
            citire_fisier()
        except:
            pass


def main2(brand,month):
    global Lista, Lista_temp, mesaj, Lista, EOF, ListaRapoarteNumeric, Lista_temp, StartPoint, EndPoint, VerificareInteger, ListaRapoarte, data
    for e in Lista:  # Indepartare linii care nu sunt necesare
        VerificareInteger = e.replace(" ", "").replace("0", "").replace(".", "").replace("\n", "").replace(",", "")
        if VerificareInteger.isdigit():
            pass
        elif "******************************************************" in e:
            pass
        elif month in e:
            pass
        elif ("DISTRICT" and "TOTAL") in e:
            pass
        elif e == (
        '---------------                                                                                                                     \n'):
            pass
        elif e == (
        '-----------                                                                                                                         \n'):
            pass
        elif "                                                                                                                                    " in e:
            pass
        elif "FOR ALL  CHECKOUT LOCATIONS AND VEHICLES" in e:
            pass
        elif "FOR ALL   CHECKIN LOCATIONS AND VEHICLES" in e:
            pass
        elif "FOR ALL  CHECKIN LOCATIONS AND VEHICLES" in e:
            pass
        elif "- MODIFIED -" in e:
            pass
        elif "                                                                                                                                     " in e:
            pass

        elif "SECURITY REPORTING SYSTEM" in e:
            pass
        else:
            Lista_temp.append(e)

    Lista = []  # reconstituire lista initiala din lista_temp
    for e in Lista_temp:
        Lista.append(e)

    for f in ListaRapoarte:  # indepartare caractere care nu sunt necesare de pe liniile care contin elementele din listaRapoarte
        i = 0
        for e in Lista:
            if f in e:
                Lista[i] = f.center(132, " ") + "\n"
            i += 1
        i = 0
    del Lista_temp

    # identificare inceput si sfarsit raport
    J = 0
    while J < (len(ListaRapoarte)):
        I = 0
        StartPoint = []
        EndPoint = []
        while I < (len(Lista)):
            if ListaRapoarte[J] in Lista[I]:
                StartPoint.append(I)
            if EOF in Lista[I]:
                EndPoint.append(I)
            I += 1
        S = StartPoint[0]
        E = EndPoint[J]
        K = StartPoint[0]
        while (S <= K <= E):
            ListaRapoarteNumeric[J].append(Lista[K])  # scriere in lista potrivita
            K += 1
        J += 1

def generare_raport(i, cap_tabel1, cap_tabel2, S, E):
    global ListaRapoarteNumeric, EOF
    s = S
    List_temp = []
    lista = []
    for e in ListaRapoarteNumeric[i]:  # Indepartare linii care nu sunt necesare
        if ListaRapoarte[i] in e:
            pass
        elif cap_tabel1 in e and cap_tabel1 != "":
            pass
        elif cap_tabel2 in e and cap_tabel2 != "":
            pass
        elif cap_tabel1 == e and cap_tabel1 == "":
            pass
        elif cap_tabel2 == e and cap_tabel2 == "":
            pass
        elif EOF in e:
            pass
        else:
            List_temp.append(e)
    ListaRapoarteNumeric[i] = []
    for e in List_temp:
        rand = ""
        while s < E:
            rand = rand + e[s]
            s += 1
        s = S
        ListaRapoarteNumeric[i].append(rand)



class CONTRACTE():
    global Lista_obiecte, nume_fisier_procesat
    cel_putin_1_procesare=False#daca se gaseste cel putin 1 contract in fisierul de brand1 sau cel de brand2, se poate executa functia de trimitere mail si stergere a fisierului CSV
    def __init__(self,RA,motiv,brand,month):
        CONTRACTE.cel_putin_1_procesare=True
        self.RA=RA
        self.motiv=motiv
        self.brand=brand
        self.month=month
        Lista_obiecte.append(self)
        fisier_contracte = open(nume_fisier_procesat, "a+")
        linie=str(self.RA.strip("E")+","+self.motiv+",\n")
        fisier_contracte.write(linie)
        fisier_contracte.close()
        query="INSERT INTO alerte_audit VALUES(default,'{0}','{1}','{2}','{3}')".format(self.RA.strip("E"),self.motiv,self.brand,self.month)
        try:
            cur.execute(query)
            db.commit()
            for row in cur.fetchall():
                print row[0]
        except(Exception) as error:
            print (error)
        


def main3(brand,data):
    numarobiecte = len(globals())
    # Procesare CASH RENTALS WITH A CHARGE MODIFIED AFTER CHECKIN
    generare_raport(0,
                    "     CHECKIN LOCATION              RA         MVA     C/I   C/I        NET     DIST RATE RATE OWAY                          OTR     ",
                    " NUMBER         NAME             NUMBER      NUMBER  AGENT DELAY       T+D   DRIVEN CODE CODE MISC FUEL INS  DEL  COL  ADJ  EXP  TOT",
                    12, 127)
    if "*** NO DETAILS FOR THIS REPORT ****" not in ListaRapoarteNumeric[0][0]:
        for e in ListaRapoarteNumeric[0]:
            globals()[numarobiecte] = CONTRACTE(e[20:30],"Cash rental with a charge modified after checkin",brand,data)
            numarobiecte += 1

            
    # Procesare RENTALS WHERE TIME AND DISTANCE HAS BEEN ADJUSTED
    generare_raport(1,
                    "     CHECKOUT LOCATION            RA         C/O      C/I        C/I      C/I          GROSS        ADJ       RSN         RATE      ",
                    " NUMBER         NAME            NUMBER      AGENT   LOCATION    AGENT     DATE          T+D        AMOUNT     CODE  MOP   CODE   TOT",
                    12, 126)
    if "*** NO DETAILS FOR THIS REPORT ****" not in ListaRapoarteNumeric[1][0]:
        for e in ListaRapoarteNumeric[1]:
            globals()[numarobiecte] = CONTRACTE(e[20:30],"Adjusted T&M",brand,data)
            numarobiecte += 1
            

    # Procesare CHECKIN DELAY ANALYSIS
    generare_raport(2,
                    "     CHECKIN LOCATION              RA        MVA      C/I      C/I     C/I        NET           CASH       DIST   RATE              ",
                    " NUMBER         NAME             NUMBER     NUMBER   AGENT     DATE   DELAY       T+D          AT C/I     DRIVEN  CODE  MOP   TOTAL ",
                    12, 124)
    if "*** NO DETAILS FOR THIS REPORT ****" not in ListaRapoarteNumeric[2][0]:
        for e in ListaRapoarteNumeric[2]:
            globals()[numarobiecte] = CONTRACTE(e[19:29],"Checkin delay analysis",brand,data)
            numarobiecte += 1


    # Procesare RENTALS_WHERE_INSURANCES_HAVE_BEEN_MODIFIED
    generare_raport(3,
                    "     CHECKIN LOCATION            RA        MVA      C/I    C/I  C/I  C/O    C/O        C/O       NET            RATE    AWD         ",
                    " NUMBER         NAME           NUMBER     NUMBER   AGENT  DELAY INS  INS  LOCATION    AGENT      T+D      MOP   CODE   NUMBER    TOT",
                    12, 126)
    if "*** NO DETAILS FOR THIS REPORT ****" not in ListaRapoarteNumeric[3][0]:
        for e in ListaRapoarteNumeric[3]:
            globals()[numarobiecte] = CONTRACTE(e[17:27],"Modified insurances",brand,data)
            numarobiecte += 1

            
    # Procesare CASH RENTALS WHERE RATE CODE MODIFIED AFTER CHECKOUT
    generare_raport(4,
                    "             CHECKOUT LOCATION                RA           MVA         C/O            NET         RATE       AWD                    ",
                    "", 13, 115)
    if "*** NO DETAILS FOR THIS REPORT ****" not in ListaRapoarteNumeric[4][0]:
        for e in ListaRapoarteNumeric[4]:
            globals()[numarobiecte] = CONTRACTE(e[29:39],"Modified rate code",brand,data)
            numarobiecte += 1


    # Procesare NO CHARGE RENTALS
    generare_raport(8,
                    "     CHECKIN LOCATION          RA        MVA     C/I   C/I                           DIST    C/O                                    ",
                    "", 5, 126)
    if "*** NO DETAILS FOR THIS REPORT ****" not in ListaRapoarteNumeric[8][0]:
        for e in ListaRapoarteNumeric[8]:
            globals()[numarobiecte] = CONTRACTE(e[23:33],"No charge rental",brand,data)
            numarobiecte += 1


    ### Procesare UNAUTHORISED DIRECT BILLING RENTALS
    ##generare_raport(9,
    ##                "    CHECKOUT LOCATION          RA         MVA     C/O        AMOUNT   DUE      C/I   DAYS RATE                      CR   C/O        ",
    ##                "", 0, 128)
    ##if "*** NO DETAILS FOR THIS REPORT ****" not in ListaRapoarteNumeric[9][0]:
    ##    for e in ListaRapoarteNumeric[9]:
    ##        globals()[numarobiecte] = CONTRACTE(e[28:38],"Unauthorised direct billing",brand,data)
    ##        numarobiecte += 1


    ### Procesare EMPLOYEE RENTALS
    ##generare_raport(10,
    ##                "    CHECKOUT LOCATION           RA        MVA      C/O     RENTAL                                          RNTL  C/I  RATE CCR",
    ##                "", 4, 126)
    ##if "*** NO DETAILS FOR THIS REPORT ****" not in ListaRapoarteNumeric[10][0]:
    ##    for e in ListaRapoarteNumeric[10]:
    ##        globals()[numarobiecte] = CONTRACTE(e[25:35],"Employee rentals",brand,data)
    ##        numarobiecte += 1

            
    ### Procesare LONG DURATION RENTALS
    ##generare_raport(12,
    ##                "          CHECKOUT LOCATION           RA        MVA      C/O      C/O       C/I        RENTAL    RNTL  RATE   DIST  CCR          ",
    ##                "", 10, 119)
    ##if "*** NO DETAILS FOR THIS REPORT ****" not in ListaRapoarteNumeric[12][0]:
    ##    for e in ListaRapoarteNumeric[12]:
    ##        globals()[numarobiecte] = CONTRACTE(e[24:34],"Long duration rentals",brand,data)
    ##        numarobiecte += 1


    ### Procesare LONG DISTANCE RENTALS
    ##generare_raport(13,
    ##                "       CHECKIN LOCATION             RA         MVA       C/I       C/I     C/I    RNTL   AVG    RATE          CCR    C/O            ",
    ##                "", 7, 121)
    ##if "*** NO DETAILS FOR THIS REPORT ****" not in ListaRapoarteNumeric[13][0]:
    ##    for e in ListaRapoarteNumeric[13]:
    ##        globals()[numarobiecte] = CONTRACTE(e[25:35],"Long distance rentals",brand,data)
    ##        numarobiecte += 1
    scriere_log(str(time.ctime()+": S-a terminat procesarea fisierului pentru "+brand+"\n"))
    


def trimitere_email():
    global nume_fisier_procesat
    fromaddr = ""
    toaddr= ""
    toccaddr= ""
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = nume_fisier_procesat
    body=" Fisierul "+nume_fisier_procesat+" a fost procesat<br>"+signature.semnatura()
    msg.attach(MIMEText(body, 'html'))

    attachment = open(nume_fisier_procesat, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % nume_fisier_procesat)
    msg.attach(part)
          
    server = smtplib.SMTP_SSL('', 465)
    server.login("", "")
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr.split(",") + toccaddr.split(","), text)
    server.quit()
    

if DB_connected == True:
    
    #Daca cumva scriptul nu a rulat intr-o zi/cateva zile la rand, cu prima ocazie cand ruleaza verifica in baza de date ce zile nu au rulat si insereaza default False pe ele
    #preluare lista de zile din baza de date
    query = "SELECT zi_calendaristica FROM zile_rulate_audit"
    lista_zile_DB=[]
    try:
        cur.execute(query)
        db.commit()
        for row in cur.fetchall():
            lista_zile_DB.append(row[0])
    except(Exception) as error:
        print (error)

    for zi in daterange(date(2020, int('01'), int('01')), date(time.localtime().tm_year, time.localtime().tm_mon, time.localtime().tm_mday)):
        if str(zi) not in lista_zile_DB:
            zi_din_an=str(time.strptime(str(zi),"%Y-%m-%d").tm_yday).rjust(3,"0")
            query = "INSERT INTO zile_rulate_audit VALUES (default, '{0}','{1}','false','{2}',default)".format(zi_din_an,str(zi),'brand1')
            try:
                cur.execute(query)
                db.commit()
            except(Exception) as error:
                print (error)
            query = "INSERT INTO zile_rulate_audit VALUES (default, '{0}','{1}','false','{2}',default)".format(zi_din_an,str(zi),'brand2')
            try:
                cur.execute(query)
                db.commit()
            except(Exception) as error:
                print (error)                


    for brand in ["brand1","brand2"]:
        if brand=="brand1":
            query = "SELECT * FROM zile_rulate_audit WHERE rulat='false' and brand='brand1'"
        else:
            query = "SELECT * FROM zile_rulate_audit WHERE rulat='false' and brand='brand2'"
            
        try:
            cur.execute(query)
            db.commit()
            for row in cur.fetchall():
                data_nerulata=row[2]
                zi_formatata = datetime.strftime(datetime.strptime(data_nerulata,"%Y-%m-%d"),"%d%b%y")
                zi_din_an = str(time.strptime(data_nerulata,"%Y-%m-%d").tm_yday).rjust(3,"0")
                an_data_nerulata = data_nerulata[0:4]
                luna_data_nerulata = str(int(data_nerulata[5:7].lstrip('0'))-1)
                month = months.get(luna_data_nerulata).strip(' ')
                if month == 'DECEMBER':
                    data = month + ' ' + str(int(an_data_nerulata) - 1)
                else:
                    data = month + ' ' + an_data_nerulata
                nume_fisier_procesat = "Raport audit - " + brand + ' ' + data + ".csv"
                if brand=="brand1":
                    Fisier='country ftp\\ro\\eprosrmc' + zi_din_an + '.zip'
                else:
                    Fisier='EBRNTDC0X1\\RO\\bud_eprosrmc' + zi_din_an + '.zip'
   
                if os.path.isfile(Fisier):
                    query = "UPDATE zile_rulate_audit SET rulat='true' WHERE zi_din_an='{0}' and brand='{1}'".format(zi_din_an,brand)
                    try:
                        cur.execute(query)
                        db.commit()
                    except(Exception) as error:
                        print (error)
                    anul_crearii_fisierului=str(time.localtime(os.path.getmtime(Fisier)).tm_year)
                    if an_data_nerulata==anul_crearii_fisierului:
                            mesaj=str(time.ctime())+": Fisierul "+Fisier+" exista.\n"
                            scriere_log(mesaj)
                            resetare_variabile(brand)
                            main1(brand)
                            main2(brand,month)
                            main3(brand,data)
                            if CONTRACTE.cel_putin_1_procesare:
                                trimitere_email()    
                                os.remove(os.path.join(os.getcwd(), nume_fisier_procesat))
                    else:
                        mesaj=str(time.ctime())+": Fisierul "+Fisier+" exista, insa nu este din anul curent.\n"
                        scriere_log(mesaj)            
                else:
                    mesaj=str(time.ctime())+": Fisierul "+Fisier+" nu exista.\n"
                    scriere_log(mesaj)
                    query = "UPDATE zile_rulate_audit SET rulat='true' WHERE zi_din_an='{0}' and brand='{1}'".format(zi_din_an,brand)
                    try:
                        cur.execute(query)
                        db.commit()
                    except(Exception) as error:
                        print (error)

        except(Exception) as error:
            print (error)


