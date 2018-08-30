module EC
end
def intro()
puts "               ********** WARNING *************

 #####################################################################
   PPPPPP  EEEEEE  CCCCCC          GGGGGG   EEEEEE   IIIIII   QQQQQQ
   PP  PP  EE      CC              GG       EE         II     QQ  QQ
   PPPPPP  EEEEE   CC              GGGGGG   EEEEE      II     QQ  QQ
   PP      EE      CC              GG  GG   EE         II     QQ  QQ
   PP      EEEEEE  CCCCCC          GGGGGG   EEEEEE   IIIIII   QQQQQQQ
  PEC LMSA at General Electric Infrastructure Queretaro
 #####################################################################
		GEIQ LMSA LOGO Created by Galo Rojo
		
  This is Version 4.3 of Surfseal Automated Table Generator Process.
  If you get an error, please call or email:
       Antonio Barajas antonio.barajas@ge.com (+52 442 456-7611).

  Comments:
    * The script that generates the spreadsheets from surfseal output
      has been modified to use the number of elements in the sheets.
    * REMEMBER THAN DURING THE GENERATION OF SPREADSHEETS IS NOT POSSIBLE
      TO OPEN A NEW EXCEL AND COPY AND PASTE FUNCTION.  IF IS REQUIRED TO
      MOVE INFORMATION USE CONTROL KEY AND MOUSE TO MAKE THE COPY.
    * The excel sheet has a limiting number of patches to postprocess
      due to microsoft restrictions.  The Max number of patches is 90 per
      table.  If you have a higher number of patches it will no prompt an
      error, just the spreadsheet will postprocess 90 tables.
    * If you found an error relative to Time, the issue is a combination
      of high number of patches and elements per patch, the best solution
      is to reduce the number of patches to postprocess per script.
    * This script run with new Office 2007 and 2010 if you have any comments
      regarding this update, please let me know.

  Program release on April 2009, Update on August 2015
"
end
def getfolder()
  folder=Dir.getwd
  return folder
end
def verifyscript(h)
  a=`ls *.rst`.downcase.split("\n")
  fnum=h[0].index("FILN")
  eff=0
  i=0
  b=[]
  h.each do |data|
    if eff<2 then
      if data.include?("FILN") then
        eff=eff+1      
      elsif data.size==1 then
        eff=eff+1
      else
				b[i]=data[fnum].gsub("surfseal_","").gsub(".uof",".rst")
				i=i+1
      end
    end
  end
  if a.sort==b.sort then
		return 1
	else
		return 0
	end
end
def createinputfile(hh)
  fend=File.new("fexit.dat","w+")
  fend<<"mkdir 04_data\n"
  fend<<"mv *.xlsx 04_data\n"
  fend<<"mv *.xls 04_data\n"
	fend<<"cp 00_surfseal_table_generator_input/surfseal_table_postprocess.xlsm surfseal_table_postprocess.xlsm\n"
  fend.close
  h=File.new("surf.lock","w+")
  h.close
  f=File.new("surfseal_out_list.dat","w+")
  hcount=0
  tnum=hh[0].index("TORQ")
  fnum=hh[0].index("FILN")  
  eff=0
  tval=0
  hh.each do |data|
    if eff<2 then
      if data.include?("FILN") then
        eff=eff+1      
      elsif data.size==1 then
        eff=eff+1
      else
        if FileTest.exist?("surfseal_inp_torque.dat") then
          tfile=File.read("surfseal_inp_torque.dat")
          if data[fnum].gsub("surfseal_","").gsub(".uof","")==tfile.downcase.split[0].gsub(".rst","") then
            tval=data[tnum].to_f
						f<<data[fnum].gsub("surfseal_","").gsub(".uof","")+" 0 0\n"
          else
            ttval=(data[tnum].to_f/tval)*100000
						if ttval==0 then
							f<<data[fnum].gsub("surfseal_","").gsub(".uof","")+" 0 0\n"
						else
							f<<data[fnum].gsub("surfseal_","").gsub(".uof","")+" "+tfile.downcase.split[0]+" 0.00"+(ttval.to_i.to_f/1000000).to_s[2..9]+"\n"
						end
          end
        else
          f<<data[fnum].gsub("surfseal_","").gsub(".uof","")+" 0 0\n"
        end    
        hcount=hcount+1
      end
    end
  end
  f.close
  g=File.new("surfseal_run.bat","w+")
  g<<"mkdir 01_resume
mkdir 02_stress
mkdir 03_mcase
while read jobname1 jobname2 tscale
do
echo $jobname1 $jobname2 $tscale
"+@Arg_1+" << log
xsurf
mcas,on
dbgen,on
rrst,${jobname1}.rst,
"
  if FileTest.exist?("surfseal_inp_torque.dat") then
    g<<"rrst,${jobname2},,,1.0,,,${tscale}
"
  end
  g<<File.read("surfseal_inp_patches.dat").upcase.gsub("SIGE","SIG1")
  g<<"fini
log
mv file18.dat surfseal_${jobname1}.out
mv file21.dat surfseal_${jobname1}.roll
mv file21.ps surfseal_${jobname1}.ps
mv file19.dat surfseal_${jobname1}.seal_geom
mv file29.dat surfseal_${jobname1}.max
mv file26.dat surfseal_${jobname1}.stress
mv file31.dat surfseal_${jobname1}.uif
mv file32.dat surfseal_${jobname1}.uof
mv *.out 01_resume
mv *.roll 01_resume
mv *.ps 01_resume
mv *.seal_geom 01_resume
mv *.max 01_resume
mv *.uif *.uof 03_mcase
mv *.stress 02_stress
rm file27.dat file30.dat
mv file*.dat 01_resume

done < surfseal_out_list.dat
rm surf.lock
"
  g.close
end

def runsurfseal()
 `bash surfseal_run.bat`
  while FileTest.exist?("surf.lock") == true
    sleep 4
  end
  `rm *.bat`
  if FileTest.exist?("surfseal_inp_torque.dat") then
    File.read("surfseal_out_list.dat").split("\n").each do |line|
      if line.split[2].to_f>0 then
        fmodif=File.read("02_stress/surfseal_"+line.split[0]+".stress")
        i=0
        af=File.new("02_stress/surfseal_"+line.split[0]+".stress","w+")
        fmodif.split("####################################################################################################################################\n").each do |section|
          if i == 0 then
            af<<section
          else
            af<<"####################################################################################################################################\n"
            af<<section.split("\n\n")[0]+"\n\n"+section.split("\n\n")[1]+"\n\n"
						val=""
            section.split("\n\n")[2].split("\n").each do|line|
              val=((line.split[7].to_f)*10-700).to_i.to_s
              val[-1,0]="."
              val=val.rjust(line.split[7].size)
              af<<line.gsub(line.split[7],val)+"\n"
            end
          end
          i = i + 1
        end
        af.close
      end
    end
  end
end
def getstressdata()
  b=Hash.new
  stressfiles = `ls -G 02_stress`.split(".stress\n")
  stressfiles.each do |stress|
    a=File.read(Fld+"/02_stress/"+stress+".stress").split("####################################################################################################################################\n")
    b[stress]=Hash.new
    n=0
    a.each do |patch|
      if n>0 then
        pnam=patch.split("\n")[0].to_s.gsub("Patch   ","Region")
        b[stress][pnam]=patch.gsub("LEFT EDGE","LEFTEDGE").gsub("RIGH EDGE","RIGHEDGE").gsub("TOP  EDGE","TOPEDGE").gsub("BOTT EDGE","BOTTEDGE").gsub("SURF CENTRD","SURFCENTRD").gsub("LEFT TOP  CORNER","LEFTTOPCORNER").gsub("RIGH TOP  CORNER","RIGHTOPCORNER").gsub("LEFT BOTT CORNER","LEFTBOTTCORNER").gsub("RIGH BOTT CORNER","RIGHBOTTCORNER").gsub("     XDIR: -","").squeeze(" ").gsub(" ",",").gsub(",\n","\n").gsub("\n,","\n").split("\n")
        b[stress][pnam][1]=patch.squeeze(" ").split("\n")[1]
        b[stress][pnam][5]=b[stress][pnam][5].gsub(/MAX,ON,SIG1\(,\d+,X,\d+\),/,"")
        b[stress][pnam][8]="3D ELEM,SIGX,SIGY,TXY,SIG1,SIG2,SIGE,TEMP,COL,ROW,LOCATION"
        i=7
        while i>=0
          if i==1 then
          elsif i== 5 then
          else
            b[stress][pnam].delete_at(i)
          end
          i=i-1
        end
      end
      n=1
    end
  end
  return b
end
def readidentifier()
  h=File.open("surfseal_inp_identifier.dat")
  hh=Hash.new
  j=0
  col=0
  tcol=0
  fcol=0
  hhh=Array.new
  h.read.upcase.split("\n").each do|line|
    if j<2 then
      if line.split.size>1 then
        if line.include? "TIME" then
          col=line.split.size
          tcol=line.split.index("TIME")
          fcol=line.split.index("FILN")
          hhh[0]=Array.new
          hhh[0]=line.split
         else 
          i=line.split[tcol].to_i
          hh[i]=Array.new
          hh[i]=line.gsub("\t"," ").split(" ",col)
          hh[i][fcol]=("SURFSEAL_"+hh[i][fcol].gsub(".RST",".UOF")).downcase
        end
      else
        j=j+1
      end
    else
      hhh[hh.size+1]=["PARM"]
      hhh[hh.size+2]=line.split
    end
  end
  i=1
  hh.keys.sort.each do|order|
    hhh[i]=Array.new
    hhh[i]=hh[order]
    i=i+1
  end
  return hhh
end
def createmcasfile(h)
  if FileTest.exist?("surfseal_inp_torque.dat") then
    File.read("surfseal_out_list.dat").split("\n").each do |line|
      if line.split[2].to_f>0 then
        fmodif=File.read("03_mcase/surfseal_"+line.split[0]+".uof")
        textmodif=fmodif.split("$ WRITE TEMPS FOR ENTIRE SURFACE MESH\n")[1].split("$ WRITE CORNER STRESSES AS FACE NODAL STRESS")[0]
        textmodif2=""
        textmodif.split("\n").each do |row|
          val=((row.split[1].to_f)*1000-70000).to_i.to_s
          val[-3,0]="."
          val=val.rjust(row.split[1].size)
          textmodif2<<row.gsub(row.split[1],val)+"\n"
        end
        af=File.new("03_mcase/surfseal_"+line.split[0]+".uof","w+")
        af<<fmodif.split("$ WRITE TEMPS FOR ENTIRE SURFACE MESH\n")[0]+"$ WRITE TEMPS FOR ENTIRE SURFACE MESH\n"+textmodif2+"$ WRITE CORNER STRESSES AS FACE NODAL STRESS"+fmodif.split("$ WRITE CORNER STRESSES AS FACE NODAL STRESS")[1]
        af.close
      end
    end
  end
  mprefix=""
  titllines=""
  titltext=""
  if FileTest.exist?("surfseal_inp_header.dat") then
    head=File.read("surfseal_inp_header.dat")
    if head.include?("MPRE") then
      mprefix=head.split("MPRE")[1].split("\n")[1]
    else
       mprefix="surfseal_mcase"
    end
    if head.include?("TITL") then
      titllines=head.split("TITL")[1].split[0]
      titltext=head.split("TITL")[1].sub("\n","\t\n").split("\t\n")[1]
    else
      titllines="1"
      titltext="Mcase for Xsurf Results"
    end
  else
    mprefix="surfseal_mcase"
    titllines="1"
    titltext="Mcase for Xsurf Results"
  end
	gg=`ls -G 03_mcase`.split("\n")[2]
  ghgh=File.new(Fld+"/03_mcase/mcase.lock","w+")
  ghgh.close
  ggg=File.new(Fld+"/03_mcase/surfseal_mcase_input.uif","w+")
  ggg<<File.open(Fld+"/03_mcase/"+gg).read
  ggg<<"\n"
  ggg<<File.open("surfseal_inp_material.dat").read
  ggg.close
  g=File.new(Fld+"/03_mcase/surfseal_mcase_generator.txt","w+")
  g<<"MPRE
"+mprefix+"
UIFN
surfseal_mcase_input.uif
TITL "+titllines+"
"+titltext+"
FTYP
UOF
FLST
"
numcase=0
cparm=0
h.each do |data|
  if data.include?("TORQ") then
    g<<(data.join(" ")+"\n").gsub("TORQ","PRM"+data.join(" ").split("PRM").size.to_s)		
  elsif data.join(" ").include?("PARM") then
    g<<(data.join(" ")+"\n")
    cparm=1
	elsif cparm==1 then
    g<<(data.join(" ")+" TORQ\n")		
  else
    g<<(data.join(" ")+"\n")
  end
  numcase=numcase+1
end
g<<"NSTR
8	1100	1200	1300
ESTR
2200	2300	2700	4300
MAXC
"
g<<(numcase+2).to_s
"
$
"
  g.close
  fg=File.new(Fld+"/mcase_run.bat","w+")
    fg<<"cd 03_mcase
"+@Arg_1+"<<idy
mcas
.set
surfseal_mcase_generator.txt
idy
rm *.dat
rm mcase.lock
"
  fg.close
end
def runmcase()
 `bash mcase_run.bat`
  while FileTest.exist?(Fld+"/03_mcase/mcase.lock") == true
    sleep 4
  end
  `rm *.bat`
end
def runmxlife()
  mprefix=""
  if FileTest.exist?("surfseal_inp_header.dat") then
    head=File.read("surfseal_inp_header.dat")
    if head.include?("MPRE") then
      mprefix=head.split("MPRE")[1].split("\n")[1]
    else
      mprefix="surfseal_mcase"
    end
  else
    mprefix="surfseal_mcase"
  end  
#-----------------SIGE RANGE
  lock=File.new(Fld+"/03_mcase/mxlife.lock","w+")
  lock.close
  mxlf=File.new(Fld+"/mxlife_run.bat","w+")
  mxlf<<"cd 03_mcase
"+@Arg_1+"<<idy
mxlife
"+mprefix+".37

patch
11
a


idy
mv f50.dat surfseal_range.dat
mv f59.dat surfseal_range_f59.out
rm f*.dat
rm *.lock
rm *.pl
"
  mxlf.close
  `bash mxlife_run.bat`
  while FileTest.exist?(Fld.gsub("\\","/")+"/03_mcase/mxlife.lock") == true
    sleep 4
  end
  `rm *.bat`  
#-----------------MIN LIFE
  lock=File.new(Fld.gsub("\\","/")+"/03_mcase/mxlife.lock","w+")
  lock.close
  mxlf=File.new(Fld.gsub("\\","/")+"/mxlife2_run.bat","w+")
  mxlf<<"cd 03_mcase
"+@Arg_1+"<<idy
mxlife
"+mprefix+".37

patch
10
a


idy
mv f50.dat surfseal_life.dat
mv f48.dat surfseal_life.uof
mv f59.dat surfseal_life_f59.out
mv f60.dat surfseal_life_tables.out
cp "+mprefix+".37 "+mprefix+"_life.37 
rm f*.dat
rm *.pl
rm *.xls
rm *.lock
"+@Arg_1+"<<ody
uofread
"+mprefix+"_life.37
surfseal_life.uof

ody
"
  mxlf.close
#~ #ABRIR CYGWIN
  IO.popen("cygwin", "w") do |io|
    io.write("cd " + Fld.gsub("\\","/") + "\n")
    io.write("bash mxlife2_run.bat"+"\n")
    io.write("exit"+ "\n")
    io.close_write
  end
  while FileTest.exist?(Fld.gsub("\\","/")+"/03_mcase/mxlife.lock") == true
    sleep 4
  end
  `rm *.bat`
end
def getsurfdata()
  filecrit=File.open("surfseal_crit_loc.xls","w+")
  filecrit<<"\tMIN LIFE\tMIN LIFE\tSIGE RANGE\tSIGE RANGE\tMAX SIG1\tMAXSIG1\n"
  filecrit<<"PATCH NAME\t2D ELEM\tLOC\t2D ELEM\tLOC\t2D ELEM\tLOC\n"
  c=Hash.new
  critmlf=Hash.new
  critser=Hash.new
  surffiles = `ls -G 03_mcase`.split("\n")
  surffiles.each do |surf|
    if surf =~/\.dat$/ then
      a=File.read(Fld+"/03_mcase/"+surf).split("\n       ENTITY  LOC")
      surfname=surf.gsub(".dat","")
      c[surfname]=Hash.new  
      n=0
      a.each do |patch|
        if n>0 then
          pnam=patch.split("\n     SURFSEAL ")[1].split("\n")[0].to_s.strip.sub(/(PATCH     )/, 'Region')
          c[surfname][pnam]=Array.new
          c[surfname][pnam]=patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[0].to_s.split("EL2D,")
          c[surfname][pnam].delete("")
          c[surfname][pnam][0]="EL2D,EL2D LOC,SIGE RANGE,MAX TIME,MAX TEMP,MAX S11,MAX S22,MAX S33,MAX S12,MAX S23,MAX S13,EL3D,EL3D LOC,MIN TIME,MIN TEMP,MIN S11,MIN S22,MIN S33,MIN S12,MIN S23,MIN S13"
          if surfname=="surfseal_life" then
            c[surfname][pnam][0]="EL2D,EL2D LOC,CALC LIFE,MIN TIME,MAX TIME,MIN TEMP,MAX TEMP,MIN STRESS,MAX STRESS,RATIO,WALKER EXP,MISSION MIX,WALKER SALT,,,EL3D,EL3D LOC,M CALC LIFE,M MIN TIME,M MAX TIME,M MIN TEMP,M MAX TEMP,M MIN STRESS,M MAX STRESS,M RATIO,M WALKER EXP,M DAMAGE"
            critmlf[pnam] = patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[1].split("EL2D,")[1].split(",")[0] + "\t" + patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[1].split("EL2D,")[1].split(",")[2]
          elsif surfname=="surfseal_range" then
              dataval=patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[2].split("EL2D,")[1].split(",")[1]
              if dataval=="CENTROID" then
                  dataval="0"
              else
                  eltmp1=patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").gsub("EDGE,,","EDGE,").split("\n,")[2].split("EL2D,")[1].split(",")[4].to_s
                  eltmp2= patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").gsub("EDGE,,","EDGE,").split("\n,")[2].split("EL2D,")[1].split(",")[6].gsub("\n","").to_s
                  elcount=patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[0].split(","+eltmp1+","+eltmp2+",").count
                  ccount=1
                  patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[0].split(","+eltmp1+","+eltmp2+",").each do |datatemp|
                    if datatemp.split("EL2D,").last.split(",")[1]!="0" and ccount<elcount then
                        dataval=datatemp.split("EL2D,").last.split(",")[1]
                    end
                    ccount=ccount+1
                  end
              end
              critser[pnam] = patch.squeeze(" ").gsub("\n "," ").gsub(" ",",").split("\n,")[2].split("EL2D,")[1].split(",")[0] + "\t" + dataval
          end
        end
        n=1
      end
    end
  end
  critmlf.each_key do |pnam|
    filecrit<<pnam<<"\t"<<critmlf[pnam]<<"\t"<<critser[pnam]<<"\n"      
  end
  filecrit.close
  return c
end
def array2excel(b, iiexcel)
  excel = WIN32OLE.new("excel.application")
  excel.DisplayAlerts = false
  excel.Interactive = false
  excel.ScreenUpdating = false
  if iiexcel==0 then
    WIN32OLE.const_load(excel, EC)
  end
  excel.Visible = false
  b.keys.each do |book|
    workbook = excel.Workbooks.Add(EC::XlWBATWorksheet)
    i=0
    b[book].keys.sort.reverse.each do |sht|
      if i>0 then
        sheet = workbook.Worksheets.Add()
        sheet.Name = sht
      else
        workbook.Sheets('Sheet1').Select
        sheet = workbook.ActiveSheet
        sheet.Name = sht
      end
      jj=1
      filedom=File.new("temporal.csv","w+")
      ktnum=3
      b[book][sht].each do |line|
        filedom<<line+"\n"
        ktnum=ktnum+1
      end
      filedom.close
      wtemporal=excel.workbooks.open(Fld+"/temporal.csv")
      wtemporal.ActiveSheet.Range("A1:AZ"+ktnum.to_s).Copy
      sheet.Paste
      wtemporal.Close(0)
      `rm temporal.csv`
      i=1
    end
    workbook.SaveAs((Fld+"/"+book+".xlsx").gsub("/","\\"))
    excel.Interactive = true
    excel.ScreenUpdating = true
    excel.ActiveWorkbook.Close(0)
  end
  excel.Quit();
end
def getsiestaver()
  if FileTest.exist?("surfseal_inp_siesta.dat") then
    @Arg_1=File.read("surfseal_inp_siesta.dat").split("\n")[0]
  else
    @Arg_1="siesta_lite"
  end
  if @Arg_1 =="siesta_lite" || @Arg_1 =="siesta_lite test"  || @Arg_1 =="siesta"  || @Arg_1 =="siesta test"  then
    return 0
  else
    return 1
  end
end
require 'win32ole'
#~ require 'watir'
Iexcel=0
intro()
Fld=getfolder()
siestafin=getsiestaver()
if siestafin == 0 then
  h=readidentifier()
  fin=verifyscript(h)
  if fin==1 then
     createinputfile(h)
     runsurfseal()
     b=getstressdata()
     array2excel(b, Iexcel)
     createmcasfile(h)
     runmcase()
     Iexcel=1
     runmxlife()
     c=getsurfdata()
     array2excel(c, Iexcel)
     `bash fexit.dat`
     `rm fexit.dat`
  else
     fend=File.new("surfseal_tables_generator_v4.3.err","w+")
    fend<<"
     ************ ERROR *************
   YOU DO NOT HAVE ALL THE FILES IN THE DIRECTORY, THEREFORE THIS SCRIPT
   WILL STOP!!!!!, IT IS POSSIBLE THAT U HAVE MORE RST FILES THAN THE
   ONES DEFINED IN THE SURFSEAL
   *********************************************************************"
    fend.close
  end
else
  fend=File.new("surfseal_tables_generator_v4.3.err","w+")
  fend<<"
	************ ERROR *************
 THE FILE 'Surfseal_inp_siesta.dat' PROVIDED DO NOT HAVE A VALID SIESTA VERSION
 PLEASE CORRECT FILE
 *********************************************************************"
  fend.close  
end