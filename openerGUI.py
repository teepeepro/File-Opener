## Commentary File Opener 
## for OBS-Commentary (Rio2016 Version)
## Version Beta 0.1
## by Taylor Powell (taylorpow@gmail.com)
## Opens, Copies and Emails most recent Commentary team's related files or directories from network drive to local desktop
##

import tkinter, re, glob, os, subprocess, sys, time, webbrowser, shutil
from tkinter import messagebox

# Global vars for venue parsing, listing
drive = 'R:\\'
venues = []
venuedir = []
rtg = 0
CompVenues = { # Globally accessible dictionary of locations 
	'R01':'R01-C1_MRC- Maracana Stadium-FB & CER',
	'R02':'02-C2_MNZ- Maracanazinho- VO', 
	'R04':'R04-C3_OLS- J. Havelange Olympic Stadium- AT',
	'R05':'R05-C4_SBD- Sambodromo- AR, MR',
	'R06':'R06-C5_ROA- Rio Olympic Arena- GA, GR, GT',
	'R07':'07-C5_AQC- Maria Lenk Aquatics Centre- DV',
	'R08':'R08-C5_AQC- Maria Lenk Aquatics Centre- SY, WP(P)',
	'R09':'R09-C6_OAS- Olympic Aquatics Stadium- SW, WP(F)',
	'R10':'R10 to R14-C7_OTN- Olympic Tennis Center- TE\R10-C7_OTN - Central Court',
	'R11':'R10 to R14-C7_OTN- Olympic Tennis Center- TE\R11-C7_OTN - Court 1',
	'R12':'R10 to R14-C7_OTN- Olympic Tennis Center- TE\R12-C7_OTN - Court 2',
	'R15':'R15-C8_ROV- Rio Olympic Velodrome- CT',
	'R16':'R16-C8_CA1- Carioca Arena 1- BK',
	'R17':'R17-C8_CA2- Carioca Arena 2- JU, WF, WG',
	'R18':'R18-C8_CA3- Carioca Arena 3- FE, TK',
	'R19':'R19-C9_FTA- Future Arena- HB',
	'R20':'R20-C10_OGC- Olympic Golf Course- GO',
	'R21':'R21-C11.1_RC2- Pavilion 2- WL',
	'R22':'R22-C11_RC3- Pavilion 3- TT',
	'R23':'R23-C11_RC4- Pavilion 4- BD',
	'R24':'R24-C11_RC6- Pavilion 6- BX',
	'R25':'R25-C12_LAG- Lagoa Stadium- RO, CF',
	'R26':'R26-C13_BVA- Beach Volleyball Arena- BV',
	'R27':'R27-C14_FTC- Fort Copacabana- CR, SWM, TR',
	'R28':'R28-C15_GLO- Marina da Gloria- SA',
	'R29':'R29-C16_PON- Pontal- TM, RW',
	'R30':'R30-C17_EQC- Olympic Equestrian Centre- ED, ES',
	'R32':'R32-C18_OSC- Olympic  Shooting Centre- SH',
        'R33':'R33-C18_OSC- Olympic Shooting Centre- SH',
	'R34':'R34-C19_OHK- Olympic Hockey Centre- HO',
	'R36':'R36-C19_YOA- Youth Arena- BK (P), PF',
	'R37':'R37-C19.1_DAQ- Deodoro Aquatics Centre- MP',
	'R38':'R38-C20_DES- Deodoro Stadium- MP, RU',
	'R39':'R39-C21_MBT- Mountain Bike Track- CM',
	'R40':'R40-C22_WWS- Whitewater Stadium- CS',
	'R41':'R41-C22_BMX- Olympic BMX Center- CB',
	'R42':'R42-C23_SPA- Itaquera Arena- FB',
	'R43':'R43-C24_MIN- Mineirao- FB',
	'R44':'R44-C25_MGA- Mane Garrincha Stadium- FB',
	'R45':'R45-C26_FNA- Fonte Nova Arena- FB',
	'R46':'R46-C27_AMZ- Amazonia Arena- FB'
}

# Functions adapted from py scripts for opening, coping and emailing selected drawing

def defineVenue (venues): # Matches r## entry from user to directory name
    for k, v in CompVenues.items():
        if k in venues:
            venuedir.append(v)
    return venuedir

def openVenueFile (drawtype, venuedir): # Open the venue file 
    go = []
    ready = []
    baddirs = ()
    print(drawtype)
    successfilecount = 0
    if drawtype == 'CCRLYT':
        print('CCRLYT selected, YAY!')
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CCRLYT_CCR Equip Layout', 'Draft')
            if os.path.exists(joineddir) == True:
                ready.append(joineddir)
            else:
                altjoineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CCRLYT_CCR Equipment Layout', 'Draft')
                ready.append(altjoineddir)
        for eachdir in ready:
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                print(directory, 'doesn´t exist in Equip or Equipment folder variations, or is listed wrong on server, so we can''t open it...')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
            print('Opening now, please wait\n', go)
            for file in go:
                subprocess.Popen(file,shell=True)
    if drawtype == 'BIOLYT':
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'BIOLYT_BIO Equip Layout', 'Draft')
            if os.path.exists(joineddir) == True:
                ready.append(joineddir)
            else:
                altjoineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'BIOLYT_BIO Equipment Layout', 'Draft')
                ready.append(altjoineddir)
        for eachdir in ready:
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                print(directory, 'doesn´t exist in Equip or Equipment folder variations, or is listed wrong on server, so we can''t open it...')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
            print('Opening now, please wait\n', go)
            for file in go:
                print(file, 'before')
                subprocess.Popen(file,shell=True)
                print(file, 'after')
    elif drawtype == 'CCRLOC':
        for directory in venuedir:  # makes source directory on network drive
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CCRLOC_CCR Location', 'Draft')
            ready.append(joineddir)
        for eachdir in ready:   # finds most recent PDF file in source directory
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                print(directory, 'Mmmmhmmm the folder that the file should be in isn\'t named correctly,\n so this program cannot find it.. grrrrr, will skip')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)   
        print('Opening now, please wait')
        for file in go: # opens file 
            subprocess.Popen(file,shell=True)
    elif drawtype == 'CPPATH':
        print('CP Path selected, yaaaassss!')
        for directory in venuedir:  # makes source directory on network drive
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CPPATH_Cable Path CPArea', 'Draft')
            ready.append(joineddir)
        for eachdir in ready:   # finds most recent PDF file in source directory
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                print(directory, 'Mmmmhmmm the folder that the file should be in isn\'t named correctly,\n so this program cannot find it.. grrrrr, will skip')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
        print('Opening now, please wait')
        for file in go: # copies file over to desktop path
            subprocess.Popen(file,shell=True)
    elif drawtype == 'CPLYT':
        print('CP Layout  selected, suuuper!')
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CPLYT_CPs Area Location', 'Draft')
            ready.append(joineddir)
        for eachdir in ready:
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                print(directory, 'Mmmmhmmm the folder that the file should be in isn\'t named correctly,\n so this program cannot find it.. grrrrr, will skip')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
        print('Opening now, please wait')
        for file in go:
            subprocess.Popen(file,shell=True)
    elif drawtype == 'CPALLOC':
        print('CP Allocation selected, Wooohoooo!')
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Departments-TC', 'Engineering', '06. Commentary', 'Rio 2016', 'Drawings', '06. CPALLOC')
            ready.append(joineddir)
        for eachdir in ready:
            for v in venues:
                newest = max(glob.iglob(v + ('*.[Pp][Dd][Ff]')), key = os.path.getctime)
                yippie = os.path.join(eachdir, newest)
                go.append(yippie)
        print('Opening now, please wait')
        for file in go:
            subprocess.Popen(file,shell=True)
    else:
          messagebox.showerror('Opps!','The file you requested isn''t in the correct location, try the old fashion way...')   
    return


def copyVenueFile(drawtype, venuedir): # Copies selected drawing file to desktop
    go = []
    ready = []
    baddirs = ()
    successfilecount = 0
    if drawtype == 'DRAWINGS':
        #print(venuedir)
        if 'R33-C18_OSC- Olympic Shooting Centre- SH' in venuedir: # This fixes the problem of the sym link of drawings for R33 being the same as R32 OSC
            venuedir.remove('R33-C18_OSC- Olympic Shooting Centre- SH')
        #print(venuedir)
        for directory in venuedir:
            ready.append(os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '2.0 Drawings', 'RIO2016', 'Released'))
        for eachdir in ready:
            os.chdir(eachdir)
            all_subdirs = [d for d in os.listdir('.') if os.path.isdir(d)]
            latest_subdir = max(all_subdirs, key=os.path.getmtime)
            print('\n' + latest_subdir + ' is directory to be copied')
            yippie = os.path.join(eachdir, latest_subdir)
            go.append(yippie)
        for everydir in go:
            userhome = os.path.expanduser('~')
            desktop = os.path.join(userhome, 'Desktop')
            os.chdir(desktop)
        for v in venues:
            ## find out how to add 3 venue letter code after _ symbol, regex (3 char after _ )
            newhomedir = v + '_Drawings_' + '(' + latest_subdir + ') on ' +  time.strftime("%d-%m-%Y") + '_' + time.strftime("%H-%M-%S")
            desktopon = os.path.join(desktop, newhomedir)
            print('Copying now, please wait')
            shutil.copytree(everydir, desktopon)
            print(latest_subdir + ' copied to ' + newhomedir + '\n')
    elif drawtype == 'CCRLYT':
        print('CCRLYT selected, YAY!')
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CCRLYT_CCR Equip Layout', 'Draft')
            if os.path.exists(joineddir) == True:
                ready.append(joineddir)
            else:
                altjoineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CCRLYT_CCR Equipment Layout', 'Draft')
                ready.append(altjoineddir)
        for eachdir in ready:
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                print(directory, 'doesn´t exist in Equip or Equipment folder variations, or is listed wrong on server, will skip')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            #print('\n' + newest + ' is file to be copied')
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
        userhome = os.path.expanduser('~')
        desktop = os.path.join(userhome, 'Desktop')
        os.chdir(desktop)
        for file in go:
            print('Copying now, please wait')
            shutil.copy(file, desktop)
            print(file + ' copied to ' + desktop + '\n')
        #print(successfilecount, 'out of', str(len(venues)), 'copied\n')
        #if len(successfilecount) != len(directory):
        #    print('The following venues were not copied because their directory doesn\'t exist (most likely the file structure is wrong):\n', baddirs)
    elif drawtype == 'CCRLOC':
        for directory in venuedir:  # makes source directory on network drive
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CCRLOC_CCR Location', 'Draft')
            ready.append(joineddir)
        for eachdir in ready:   # finds most recent PDF file in source directory
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                print(directory, 'Mmmmhmmm the folder isn\'t named correctly,\n so this program cannot find it.. grrrrr, will skip')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            print('\n' + newest + ' is file to be copied')
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
        userhome = os.path.expanduser('~') # gets desktop path and changes to it
        desktop = os.path.join(userhome, 'Desktop')
        os.chdir(desktop)   
        for file in go: # copies file over to desktop path
            print('Copying now, please wait')
            shutil.copy(file, desktop)
            print(file + ' copied to desktop\n')
        print(successfilecount, 'out of', str(len(venues)), 'copied\n')
        #if len(successfilecount) != len(directory):
        #    print('The following venues were not copied because their directory doesn\'t exist (most likely the file structure is wrong):\n', baddirs)
    elif drawtype == 'CPPATH':
        print('CP Path selected, yaaaassss!')
        for directory in venuedir:  # makes source directory on network drive
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CPPATH_Cable Path CPArea', 'Draft')
            ready.append(joineddir)
        for eachdir in ready:   # finds most recent PDF file in source directory
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                print(directory, 'Mmmmhmmm the folder isn\'t named correctly,\n so this program cannot find it.. grrrrr, will skip')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            print('\n' + newest + ' is file to be copied')
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
        userhome = os.path.expanduser('~')  # gets desktop path and changes to it
        desktop = os.path.join(userhome, 'Desktop')
        os.chdir(desktop)
        for file in go: # copies file over to desktop path
            print('Copying now, please wait')
            shutil.copy(file, desktop)
            print(file + ' copied to desktop\n')
        print(successfilecount, 'out of', str(len(venues)), 'copied\n')
        #if len(successfilecount) != len(directory):
        #    print('The following venues were not copied because their directory doesn\'t exist (most likely the file structure is wrong):\n', baddirs)
    elif drawtype == 'CPLYT':
        print('CP Layout  selected, suuuper!')
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CPLYT_CPs Area Location', 'Draft')
            ready.append(joineddir)
        for eachdir in ready:
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                print(directory, 'Mmmmhmmm the folder isn\'t named correctly,\n so this program cannot find it.. grrrrr, will skip')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            print('\n' + newest + ' is file to be copied')
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
        userhome = os.path.expanduser('~')
        desktop = os.path.join(userhome, 'Desktop')
        os.chdir(desktop)
        for file in go:
            print('Copying now, please wait')
            shutil.copy(file, desktop)
            print(file + ' copied to desktop\n')
        print(successfilecount, 'out of', str(len(venues)), 'copied\n')
        #if len(successfilecount) != len(directory):
        #    print('The following venues were not copied because their directory doesn\'t exist (most likely the file structure is wrong):\n', baddirs)
    elif drawtype == 'CPALLOC':
        print('CP Allocation selected, Wooohoooo!')
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Departments-TC', 'Engineering', '06. Commentary', 'Rio 2016', 'Drawings', '06. CPALLOC')
            ready.append(joineddir)
        for eachdir in ready:
            for v in venues:
                newest = max(glob.iglob(v + ('*.[Pp][Dd][Ff]')), key = os.path.getctime)
                print('\n' + newest + ' is file to be copied')
                yippie = os.path.join(eachdir, newest)
                go.append(yippie)
        userhome = os.path.expanduser('~')
        desktop = os.path.join(userhome, 'Desktop')
        os.chdir(desktop)
        for file in go:
            print('Copying now, please wait')
            shutil.copy(file, desktop)
            print(file + ' copied to desktop\n')
        print(str(len(go)), 'out of', str(len(venues)), 'copied\n')
        #if len(go) != len(directory):
        #    print('The following venues were not copied because their directory doesn\'t exist (most likely the file structure is wrong):\n', baddirs)
    else:
          print('Invalid selection, select from this list: ' + str(typeoptions))     
    return


def emailVenueFile (drawtype, venuedir): # Emails the venue file as attachment (using outlook only!)
    go = []
    ready = []
    baddirs = ()
    successfilecount = 0
    if drawtype == 'CCRLYT':
        print('CCRLYT selected, YAY!')
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CCRLYT_CCR Equip Layout', 'Draft')
            if os.path.exists(joineddir) == True:
                ready.append(joineddir)
            else:
                altjoineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CCRLYT_CCR Equipment Layout', 'Draft')
                ready.append(altjoineddir)
        for eachdir in ready:
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                messagebox.showerror('Opps!','File doesn´t exist in Equip or Equipment folder variations, or is listed wrong on server, so we can''t open it...')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
            print('Opening now, please wait\n', go)
            for file in go:
                webbrowser.open("mailto:")
    if drawtype == 'BIOLYT':
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'BIOLYT_BIO Equip Layout', 'Draft')
            if os.path.exists(joineddir) == True:
                ready.append(joineddir)
            else:
                altjoineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'BIOLYT_BIO Equipment Layout', 'Draft')
                ready.append(altjoineddir)
        for eachdir in ready:
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                messagebox.showerror('Opps!','File doesn´t exist in Equip or Equipment folder variations, or is listed wrong on server, so we can''t open it...')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
            print('Opening now, please wait\n', go)
            for file in go:
                webbrowser.open("mailto:")
    elif drawtype == 'CCRLOC':
        for directory in venuedir:  
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CCRLOC_CCR Location', 'Draft')
            ready.append(joineddir)
        for eachdir in ready:   
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                messagebox.showerror('Opps!','Mmmmhmmm the folder that the file should be in isn\'t named correctly,\n so this program cannot find it.. grrrrr, will skip')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)   
        print('Opening now, please wait')
        for file in go: 
            webbrowser.open("mailto:")
    elif drawtype == 'CPPATH':
        print('CP Path selected, yaaaassss!')
        for directory in venuedir:  
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CPPATH_Cable Path CPArea', 'Draft')
            ready.append(joineddir)
        for eachdir in ready:   
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                messagebox.showerror('Opps!', 'Mmmmhmmm the folder that the file should be in isn\'t named correctly,\n so this program cannot find it.. grrrrr, will skip')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
        print('Opening now, please wait')
        for file in go: 
            webbrowser.open("mailto:")
    elif drawtype == 'CPLYT':
        print('CP Layout  selected, suuuper!')
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Games', '2016 Rio', 'Olympics', '2. Competition Venues', directory, '15.0 Commentary Systems', 'CPLYT_CPs Area Location', 'Draft')
            ready.append(joineddir)
        for eachdir in ready:
            try:
                os.chdir(eachdir)
                successfilecount += 1
            except FileNotFoundError:
                messagebox.showerror('Opps!','Mmmmhmmm the folder that the file should be in isn\'t named correctly,\n so this program cannot find it.. grrrrr, will skip')
                successfilecount -= 1
                baddirs = directory
                continue
            newest = max(glob.iglob('*.[Pp][Dd][Ff]'), key = os.path.getctime)
            yippie = os.path.join(eachdir, newest)
            go.append(yippie)
        print('Opening now, please wait')
        for file in go:
            webbrowser.open("mailto:")
    elif drawtype == 'CPALLOC':
        print('CP Allocation selected, Wooohoooo!')
        for directory in venuedir:
            joineddir = os.path.join(drive, 'Departments-TC', 'Engineering', '06. Commentary', 'Rio 2016', 'Drawings', '06. CPALLOC')
            ready.append(joineddir)
        for eachdir in ready:
            for v in venues:
                newest = max(glob.iglob(v + ('*.[Pp][Dd][Ff]')), key = os.path.getctime)
                yippie = os.path.join(eachdir, newest)
                go.append(yippie)
        print('Opening now, please wait')
        for file in go:
             webbrowser.open("mailto:")
    else:
          messagebox.showerror('Opps!','Say what??!?!?! Um, email isn''t working yet')     
    return

# Here is the GUI portion:

class app_tk(tkinter.Tk):
    def __init__(self,parent):
        tkinter.Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()
        
    def initialize(self): # Buttons, Entry box, text box
        self.grid()

        self.entryVariable = tkinter.StringVar()
        self.entry = tkinter.Entry(self,textvariable=self.entryVariable)
        self.entry.grid(column=0,row=0,sticky='EW')
        self.entry.bind("<Return>", self.OnPressEnter)
        self.entryVariable.set(u"Enter venue/drawing here.")

        button = tkinter.Button(self,text=u"Open",
                                command=self.OnButtonClickOpen)
        button.grid(column=2,row=0,sticky='N')
        button2 = tkinter.Button(self,text=u"Copy to Desktop",
                                command=self.OnButtonClickCopy)
        button2.grid(column=2,row=1,sticky='N')
        button3 = tkinter.Button(self,text=u"Email",
                                command=self.OnButtonClickEmail)
        button3.grid(column=2,row=2,sticky='N')

        self.labelVariable = tkinter.StringVar()
        label = tkinter.Label(self,textvariable=self.labelVariable,
                              anchor="w",fg="white",bg="black")
        label.grid(column=0,row=1,columnspan=2,sticky='EW')
        self.labelVariable.set(u'''Enter Venue(s):
        R## and
        CCRLYT  CCRLOC  CPPATH
        CPLYT  CPALLOC or BIOLYT
        ''')

        self.grid_columnconfigure(0,weight=2)
        self.resizable(False,False)
        self.update()
        self.geometry(self.geometry())
        self.entry.focus_set()
        self.entry.selection_range(0, tkinter.END)


    def inputClean(self): # Main RegEx splicer to get r## and drawing type out of entry
        inputEntry = self.entryVariable.get()
        inputEntry = inputEntry.upper()
        venueReGex = re.compile('R\\d\\d')
        coo = venueReGex.findall(inputEntry)
        defineVenue(coo)
        inputEntry = inputEntry.upper()
        drawingReGex = re.compile('CCRLYT|CCRLOC|CPPATH|CPLYT|CPALLOC|BIOLYT')
        rad = drawingReGex.search(inputEntry)
        global rtg
        try:
            rtg = rad.group()
        except AttributeError:
            messagebox.showerror('Incorrect Entry', 'Check your entry and try again, because it is WRONG!')
        return rtg
    

    def OnButtonClickOpen(self):
        inputEntry = self.entryVariable.get()
        self.inputClean()
        openVenueFile(rtg, venuedir)
        self.entry.focus_set()
        self.entry.selection_range(0, tkinter.END)

    def OnButtonClickCopy(self):
        inputEntry = self.entryVariable.get()
        self.inputClean()
        copyVenueFile(rtg, venuedir)
        self.entry.focus_set()
        self.entry.selection_range(0, tkinter.END)
        
    def OnButtonClickEmail(self):
        inputEntry = self.entryVariable.get()
        self.inputClean()
        emailVenueFile(rtg, venuedir)
        self.entry.focus_set()
        self.entry.selection_range(0, tkinter.END)
        
    def OnPressEnter(self,event):
        inputEntry = self.entryVariable.get()
        self.inputClean()
        openVenueFile(rtg, venuedir)
        self.entry.focus_set()
        self.entry.selection_range(0, tkinter.END)
        
if __name__ == "__main__":
    app = app_tk(None)
    app.title('Commentary Opener (beta)')
    app.mainloop()
    
