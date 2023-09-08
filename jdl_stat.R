setwd('/Users/yuanyuanji/Desktop')
library("readxl")
library("ggplot2")
semester<-c(paste('F',15:21,sep=''),paste('S',16:22,sep=''))
for (s in semester){
assign(s,read_excel("CO-OP.xlsx",sheet=s))
}
final<-data.frame()
for (s in c('F15','S16','F16','S17')){
  data<-get(s)
  for (i in 2:nrow(data)){
    First_Name<-strsplit(as.character(data[i,1]),split=',')[[1]][2]
    Last_Name<-strsplit(as.character(data[i,1]),split=',')[[1]][1]
    Major<-as.character(data[i,2])
    Semester<-s
    Company_List<-rownames(na.omit(as.data.frame(t(data[i,-c(1:2)]))))
    Status_List<-na.omit(as.data.frame(t(data[i,-c(1:2)])))[,1]
    if (length(Company_List)!=0){
    for (co in 1:length(Company_List)){
        Company<-Company_List[co]
        Status<-Status_List[co]
        if (co!=1){
          Semester=First_Name=Last_Name=Major=''
        }
        final<-rbind(final,cbind(Semester,First_Name,Last_Name,Major,Company,Status))
      }
    }}
#    table(na.omit(as.data.frame(t(data[i,-c(1:2)]))))
#    Applied<-sum(!is.na(data[i,-c(1:2)]))
#    Accepted<-sum(t(data[i,-c(1:2)])%in%c('Accepted','Offer Extended'))
    #Applied<-
#    final<-rbind(final,cbind(Semester,First_Name,Last_Name,Major,Applied,Accepted))
}
for(s in c('F17','S18','F18','S19','F19','S20','F20','S21','F21') ){
  data<-get(s)
  for (i in 2:nrow(data)){
    str<-strsplit(as.character(data[i,1]),split=' ')[[1]]
    len<-length(str)
    if (len==2){
      First_Name<-str[1]
      Last_Name<-str[2]
    } else {
      First_Name<-paste(str[-len],collapse=' ')
      Last_Name<-str[len]
    }
    Major<-as.character(data[i,2])
    Semester<-s
    Company_List<-rownames(na.omit(as.data.frame(t(data[i,-c(1:2)]))))
    Status_List<-na.omit(as.data.frame(t(data[i,-c(1:2)])))[,1]
    if (length(Company_List)!=0){
      for (co in 1:length(Company_List)){
        Company<-Company_List[co]
        Status<-Status_List[co]
        if (co!=1){
          Semester=First_Name=Last_Name=Major=''
        }
        final<-rbind(final,cbind(Semester,First_Name,Last_Name,Major,Company,Status))
      }
    }}
  #  Applied<-sum(!is.na(data[i,-c(1:2)]))
 #   Accepted<-sum(t(data[i,-c(1:2)])%in%c('Accepted','Offer Extended'))
}

for(s in c('S22') ){
  data<-get(s)
  for (i in 2:nrow(data)){
    First_Name<-as.character(data[i,1])
    Last_Name<-as.character(data[i,2])
    Major<-as.character(data[i,3])
    Semester<-s
    Company_List<-rownames(na.omit(as.data.frame(t(data[i,-c(1:1)]))))
    Status_List<-na.omit(as.data.frame(t(data[i,-c(1:1)])))[,1]
    if (length(Company_List)!=0){
      for (co in 1:length(Company_List)){
        Company<-Company_List[co]
        Status<-Status_List[co]
        if (co!=1){
          Semester=First_Name=Last_Name=Major=''
        }
        final<-rbind(final,cbind(Semester,First_Name,Last_Name,Major,Company,Status))
      }
    }}
#    Applied<-sum(!is.na(data[i,-c(1:1)]))
#    Accepted<-sum(t(data[i,-c(1:1)])%in%c('Accepted','Offer Extended'))
#    final<-rbind(final,cbind(Semester,First_Name,Last_Name,Major,Applied,Accepted))
#  }
}

final[final$Major%in%c("Civil Eng","Civil Engnieering","Civil Engineering","Civil"),]$Major="Civil Engineering"
final[final$Major%in%c("Computer Eng","ComputerEng","Computer Engineering"),]$Major="Computer Engineering"
final[final$Major%in%c("Computer Sciences","Computer Science","ComputerSci"),]$Major="Computer Science"
final[final$Major%in%c("Electrical Eng","ElectricalEng","Electrical Engineering","Electrical"),]$Major="Electrical Engineering"
final[final$Major%in%c("Eng Management","EngManagement","Engineering Management"),]$Major="Engineering Management"
final[final$Major%in%c("Mechanical Eng","Mechanical Engineering","MechanicalEng","Mechanical"),]$Major="Mechanical Engineering"
final[final$Major%in%c("Eng Physics","Enginereering Physics","EngPhysics","Engineering Physics"),]$Major="Engineering Physics"

write.csv(final,'CO-OP-STAT.csv')
write.csv(final,'CO-OP-STAT_2.csv')

##analysis
# number of students in each union who applied
final$AY<-'AY 15-16'
final[final$Semester%in%c('F16','S17'),]$AY<-'AY 16-17'
final[final$Semester%in%c('F17','S18'),]$AY<-'AY 17-18'
final[final$Semester%in%c('F18','S19'),]$AY<-'AY 18-19'
final[final$Semester%in%c('F19','S20'),]$AY<-'AY 19-20'
final[final$Semester%in%c('F20','S21'),]$AY<-'AY 20-21'
final[final$Semester%in%c('F21','S22'),]$AY<-'AY 21-22'

