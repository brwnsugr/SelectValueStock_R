setwd("/Users/brownsugar/Documents/R_Practice/Korea Equity/Korea Equity_20160708")
options(java.parameters = "-Xmx2g" )
install.packages("readxl")
library(readxl)
install.packages("xlsx")
library(xlsx)
install.packages("XLConnect")
library(XLConnect)


  rawpe<-read_excel("dart_equity_pepb_(m).xlsx", sheet="pe",FALSE) # read sheet pe (no raw&column name)
  symbol<-rawpe[-1,1]
  name<-rawpe[-1,2]
  rawdate<-rawpe[1,-1:-3] 
  rawdate<-as.numeric(rawdate) # converts data.frame to numeric
  date<-as.Date(rawdate,origin="1899-12-30") # converts excel dates to R
  date<-as.character(date) # converts date to character
  
  
  ## rank pe ##
  pe<-read_excel("dart_equity_pepb_(m).xlsx", sheet="pe",TRUE) # read sheet pe 
  pe<-pe[,-1:-3]
  pe[is.na(pe)]=100
  dimv<-dim(pe) # set dimension of valuation factor pe&pb
  
  pem<-matrix(,dimv[1],dimv[2]) # processing pe data
  for(i in 1:dimv[1])
    for(j in 1:dimv[2]){
      if(pe[i,j]<=0 || pe[i,j]=="N/A"){pem[i,j]<-100} else{pem[i,j]<-pe[i,j]}
    }
  
  pem2<-as.numeric(pem)
  pem3<-matrix(pem2,dimv[1],dimv[2])
  
  rpe<-matrix(,dimv[1],dimv[2],,list(symbol,date)) # rank pe
  for(j in 1:dimv[2]){
    rpe[,j]<-rank(pem3[,j],TRUE,"average")
  }
  
  
  ## rank pb ##
  pb<-read_excel("dart_equity_pepb_(m).xlsx", sheet="pb",TRUE) # read sheet pe 
  pb<-pb[,-1:-3]
  pb[is.na(pb)]=100
  
  pbm<-matrix(,dimv[1],dimv[2]) # processing pe data
  for(i in 1:dimv[1])
    for(j in 1:dimv[2]){
      if(pb[i,j]<=0 || pb[i,j]=="N/A"){pbm[i,j]<-100} else{pbm[i,j]<-pb[i,j]}
    }
  
  pbm2<-as.numeric(pbm)
  pbm3<-matrix(pbm2,dimv[1],dimv[2])
  
  rpb<-matrix(,dimv[1],dimv[2],,list(symbol,date)) # rank pe
  for(j in 1:dimv[2]){
    rpb[,j]<-rank(pbm3[,j],TRUE,"average")
  }
  
  
  ## LPP ##
  rawreturn<-read_excel("krx_equity_change_(m).xlsx", sheet="return",TRUE) # read sheet return 
  
  matchv<-match(symbol,rawreturn[,1])
  #matchv<-NULL # valuation ticker ????��?? return ticker ��ġ
  #for(i in 1:dimv[1]){
  #  matchv[i]<-which(rawreturn[,1]==symbol[i])
  #}
  
  return<-rawreturn[,-1:-3]
  
  returnmatch<-matrix(,dimv[1],dimv[2])
  for(i in 1:dimv[1])
    for(j in 1:dimv[2]){
      returnmatch[i,j]<-return[matchv[i],j]
    }
  returnmatch[is.na(returnmatch)]=0
  
  m<-matrix(,dimv[1],dimv[2])
  for(i in 1:dimv[1])
    for(j in 1:dimv[2]){
      m[i,j]<-mean(returnmatch[i,1:j])
    }
  m[is.na(m)]=0
  
  s<-matrix(,dimv[1],dimv[2])
  for(i in 1:dimv[1])
    for(j in 1:dimv[2]){
      s[i,j]<-sd(returnmatch[i,1:j])
    }
  s[is.na(s)]=0
  
  p<-matrix(,dimv[1],dimv[2])
  for(i in 1:dimv[1])
    for(j in 1:dimv[2]){
      p[i,j]<-pnorm(returnmatch[i,j],m[i,j],s[i,j])
    }
  
  LPP<-matrix(,dimv[1],dimv[2],,list(symbol,date))
  for(i in 1:dimv[1])
    for(j in 1:dimv[2]){
      if(p[i,j]<0.1){LPP[i,j]<-1.1} else{LPP[i,j]<-1}
    }
  
  
  ## pepb rank * LPP ##
  
  # Ticker
  
  rawticker<-read_excel("ticker.xlsx",sheet=1,FALSE) # ticker
  ticker<-rawticker[-1,]
  dimt<-dim(ticker)
  tickers<-as.character(ticker[,1])
  names<-as.character(ticker[,2])
  
  # matcht1<-NULL
  # for(i in 1:dimt[1]){
  #   matcht1[i]<-which(ticker[i,1]==symbol)
  #   }
  
  matcht1<-match(ticker[,1],symbol) # returns ticker's position of its arg in its svecond
  #match(symbol,ticker[,1])

  # Volume

  rawvolume<-read_excel("dart_equity_volume_(m).xlsx",sheet=1,FALSE) # volume
  volumet<-rawvolume[-1,1]
  rawvolumed<-rawvolume[1,-1:-3]
  volumed<-as.numeric(rawvolumed)
  volumedate<-as.Date(volumed,origin="1899-12-30")
  volumedate<-as.character(volumedate)
 
  volume<-rawvolume[-1,-1:-3]

  matchv<-match(ticker[,1],volumet)
  matchd<-match(date,volumedate)
  which(matchd==1)

  volumem<-matrix(,dimt[1],dimv[2],,list(tickers,date)) # volume matrix match ticker & date
  for(i in 1:dimt[1])
    for(j in which(matchd==1):dimv[2]){
      volumem[i,j]<-volume[matchv[i],j-(which(matchd==1)-1)]
    }
  volumem[is.na(volumem)]=0

  volumefilter<-matrix(,dimt[1],dimv[2],,list(tickers,date))
  for(i in 1:dimt[1])
    for(j in 1:dimv[2]){
      if(volumem[i,j]>300000000){volumefilter[i,j]<-1} else{volumefilter[i,j]<-0}
    }

  rpepb<-matrix(,dimv[1],dimv[2])
  for(i in 1:dimv[1])
    for(j in 1:dimv[2]){
      rpepb[i,j]<-LPP[i,j]*(0.5*rpe[i,j]+0.5*rpb[i,j])
  }

  valuation<-matrix(,dimt[1],dimv[2]) # ticker ????��?? valuation matrix
  for(i in 1:dimt[1])
    for(j in 1:dimv[2]){
      valuation[i,j]<-rpepb[matcht1[i],j]
    }
  
  valuationfilter<-matrix(,dimt[1],dimv[2]) # Volume Filter
  for(i in 1:dimt[1])
    for(j in 1:dimv[2]){
      if(volumefilter[i,j]==1){valuationfilter[i,j]<-valuation[i,j]} else {1000}
    }

  valuation2<-matrix(,dimt[1],dimv[2],,list(tickers,date))
  for(j in 1:dimv[2]){
    valuation2[,j]<-rank(valuation[,j],TRUE,"first")
  }
  
  selectvalue<-matrix(,dimt[1],dimv[2],,list(tickers,date)) # select 40 tickers
  for(i in 1:dimt[1])
    for(j in 1:dimv[2]){
      selectvalue[i,j]<-if(valuation2[i,j]<=40){1} else{0}
    }
  
  write.csv(selectvalue,"selectvalue.csv")
  
  selectvalue2<-matrix(,40,dimv[2]) # select 40 tickers
  for(i in 1:dimt[1])
    for(j in 1:dimv[2])
      for(k in 1:40)
        if(valuation2[i,j]==k){selectvalue2[k,j]<-ticker[i,2]}
  
  selectvalue2<-rbind(date,selectvalue2)
  
  ## CROGIC ##
  rawcrogic<-read_excel("dart_equity_crogic_(q).xlsx", sheet="crgic q",FALSE) # read sheet crogic (no raw&column name)
  symbol2<-rawcrogic[-1,1]
  name2<-rawcrogic[-1,2]
  rawdate2<-rawcrogic[1,-1:-3] 
  rawdate2<-as.numeric(rawdate2) # converts data.frame to numeric
  date2<-as.Date(rawdate2,origin="1899-12-30") # converts excel dates to R
  date2<-as.character(date2) # converts date to character
  
  ## rank crogic ##
  crogic<-read_excel("dart_equity_crogic_(q).xlsx", sheet="crgic q",TRUE) # read sheet crogic
  crogic<-crogic[,-1:-3]
  crogic[is.na(crogic)]=-100
  dimc<-dim(crogic) # set dimension of crogic q
  
  crogicm<-matrix(,dimc[1],dimv[2])
  for(i in 1:dimc[1])
    for(j in 1:(dimc[2]-1)){
      h<-j*3
      crogicm[i,(h-2+6):(h+6)]<-crogic[i,j]
    }
  crogicm[is.na(crogicm)]=-100
  
  dimcm<-dim(crogicm)
  
  crogicr<-matrix(,dimcm[1],dimcm[2])
  for(i in 1:dimcm[1])
    for(j in 1:dimcm[2]){
      crogicr[,j]<-rank(-crogicm[,j],TRUE,"average")
    }
  
  matchc<-match(symbol2,symbol) # returns symbol2's position of its arg in its second
  
  crogicr2<-matrix(,dimcm[1],dimcm[2])
  for(i in 1:dimcm[1])
    for(j in 1:dimcm[2]){
      crogicr2[i,j]<-crogicr[i,j]*LPP[matchc[i],j]
    }
  
  matcht2<-match(ticker[,1],symbol2)
  
  fundamental<-matrix(,dimt[1],dimcm[2])
  for(i in 1:dimt[1])
    for(j in 1:dimcm[2]){
      fundamental[i,j]<-crogicr2[matcht2[i],j]
    }
  
  fundamentalfilter<-matrix(,dimt[1],dimcm[2]) # Volume Filter
  for(i in 1:dimt[1])
    for(j in 1:dimcm[2]){
      if(volumefilter[i,j]==1){fundamentalfilter[i,j]<-fundamental[i,j]} else {1000}
  }

  fundamental2<-matrix(,dimt[1],dimcm[2],,list(tickers,date))
  for(j in 1:dimcm[2]){
    fundamental2[,j]<-rank(fundamental[,j],TRUE,"first")
  }
  
  selectcrogic<-matrix(,dimt[1],dimcm[2],,list(tickers,date))
  for(i in 1:dimt[1])
    for(j in 1:dimcm[2]){
      selectcrogic[i,j]<-if(fundamental2[i,j]<=40){1} else{0}
    }
  
  write.csv(selectcrogic,"selectcrogic.csv")
  
  selectcrogic2<-matrix(,40,dimcm[2])
  for(i in 1:dimt[1])
    for(j in 1:dimcm[2])
      for(k in 1:40)
        if(fundamental2[i,j]==k){selectcrogic2[k,j]<-ticker[i,2]}
  
  ## Set Data with summation of selectcrgoic and selectvalue
  selectboth<-selectcrogic+selectvalue
  
  # how many 1s, 2s and 0s in selectboth(to calculate)
  a<-0
  for(i in 1:790)
    for(j in 1:198)
      if(selectboth[i,j]==1){a<-a+1}
  
  matchs<-match(returnmatch2[,1],ticker[,1])
  
  
  
