###############################################################################
# Script: cell2cell_Part2_xls_logistic.R
# Copyright (c) 2019 by Alan Montgomery. Distributed using license CC BY-NC 4.0
# To view this license see https://creativecommons.org/licenses/by-nc/4.0/
#
# creates a spreadsheet "cell2cell_results.csv" using your 'best' logistic model
# from Part1 to predict the churn and LTV, which can be used to predict how
# consumers will respond to potential offers or scenarios
# (e.g., give everyone a free phone)
#
# also to help reduce the complexity of the problem it outputs results only
# for a selected set of customers (instead of all 40,000)
# The choice of customers is driven by the results of a cluster analysis
# which creates 20 segments of customers whose probability of churn is in
# the highest quartile (e.g., look at those customers most likely to churn)
###############################################################################


###############################################################################
### setup
###############################################################################

# load packages
if (!require(rstudioapi)) {install.packages("rstudioapi"); library(rstudioapi)}
if (!require(plyr)) {install.packages("plyr"); library(plyr)}
if (!require(stringr)) {install.packages("stringr"); library(stringr)}
if (!require(psych)) {install.packages("psych"); library(psych)}
if (!require(lattice)) {install.packages("lattice"); library(lattice)}
if (!require(proxy)) {install.packages("proxy"); library(proxy)}
if (!require(openxlsx)) {install.packages("openxlsx"); library(openxlsx)}

# create vector with excel column names for vec and using vec as names
excelcol = function(vec) {
  n=length(vec)  # compute length, assumes that n<=702
  letvec=as.vector(strsplit(LETTERS,""))  # create vector of letters
  colxls=as.vector(sapply(c("",letvec),function(x) {paste0(x,letvec)}))  # expand all column names
  colxls=colxls[1:n]  # only use first n
  names(colxls)=vec
  return(colxls)
}

# create worksheet for openxlsx using a data frame
# can reference columns using given names, and specify formats and comments
# dtable is the data frame to convert
# formula contains a vector of the Excel formulas that corresponds with the column names in dtable to replace
excelformula = function(dtable,xformula) {

  # create vector with appropriate excel column names for data
  xname=excelcol(colnames(dtable))
  
  # extract formulas
  xsplit=strsplit(xformula,"=")
  xcol=do.call(rbind,xsplit)[,1]
  xformula=do.call(rbind,xsplit)[,2]
  
  # replace variable names from conditions and rules with appropriate Excel variable
  for (i in 1:nrow(dtable)) {
    # append condcol with row number+1 (needs to be incremented to adjust for row with names)
    augxrow=paste0(xname,i+1)
    names(augxrow)=paste0("\\b",names(xname),"\\b")  # use \\b to only replace words
    # need to replace variable names in formulas with appropriate excel column
    dtable[i,xcol]=str_replace_all(xformula,augxrow)
  }
  
  # mark the columns as formula's for openxlsx
  for (i in xcol) {
    class(dtable[,i])=c(class(dtable[,i]),"formula")
  }

  # return the new table  
  return(dtable)
}



###############################################################################
### read in the data and prepare the dataset for analysis
###############################################################################

# set to working directory of script (assumes data in same directory as script)
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))  # only works in Rstudio scripts
# alternatively set the working directory manually
#setwd("~/Documents/class/marketing analytics/cases/cell2cell/data") #!! set to your directory

# import dataset from file (change the directory to where your data is stored)
#cell2cell=read.csv("cell2cell_data.csv")
cell2cell=read.csv("cell2cell_full.csv")  # includes a prediction sample that is not oversampled
cell2celldoc=read.csv("cell2cell_doc.csv",as.is=TRUE)  # just read in as strings not factors
cell2cellinfo=cell2celldoc$description  # create a separate vector with just the variable description
rownames(cell2celldoc)=cell2celldoc$variable   # add names so we can reference like cell2celldoc["Revenue",]
names(cell2cellinfo)=cell2celldoc$variable  # add names so we can reference like cell2cellinfo["Revenue]

# set the random number seed so the samples will be the same if regenerated
set.seed(1248765792)

# prepare new values
trainsample=(cell2cell$Sample==1)
validsample=(cell2cell$Sample==2)
predsample=(cell2cell$Sample==3)
testsample=(cell2cell$Sample==0)
plotsample=sample(1:40000,200)   # sample the first 40,000 observations (just oversampled ones)

# remove sample from the cell2cell set, since we have the sample variables
cell2cell$Sample=NULL

# recode the location so that we only keep the first 3 characters of the region
# and only remember the areas with more than 800 individuals, otherwise set region to OTH for other
newcsa=strtrim(cell2cell$Csa,3)  # get the MSA which is the first 3 characters
csasize=table(newcsa)  # count number of times MSA occurs
csasizeorig=rownames(csasize)  # save the city names which are in rownames of our table
csasizename=csasizeorig  # create a copy
csasizename[csasize<=800]="OTH"  # replace the city code to other for those with fewer than 800 customers
# overwrite the original Sca variable with the newly recoded variable using mapvalues
cell2cell$Csa=factor(mapvalues(newcsa,csasizeorig,csasizename))  # overwrites original Csa

# create a missing variable for age1 and age2
cell2cell$Age1[cell2cell$Age1==0]=NA  # replace zero ages with missing value
cell2cell$Age2[cell2cell$Age2==0]=NA  # replace zero ages with missing value
cell2cell$Age1miss=ifelse(is.na(cell2cell$Age1),1,0)  # create indicator for missing ages
cell2cell$Age2miss=ifelse(is.na(cell2cell$Age2),1,0)  # create indicator for missing ages

# replace missing values with means
nvarlist = sapply(cell2cell,is.numeric)  # get a list of numeric variables
NA2mean <- function(x) replace(x, is.na(x), mean(x, na.rm = TRUE))  # define a function to replace NA with means
cell2cell[nvarlist] = lapply(cell2cell[nvarlist], NA2mean)  # lapply performs NA2mean for each columns



###############################################################################
### estimate your logistic regression model (! only logistic allowed !)
### only estimates the model, does not do any diagnostics or statistics
### you must save your model to "mymodel" for the rest of the script
### !! use your best model from part 1
###############################################################################

# !! just uncomment one model OR paste your "mymodel=" below and comment out the others !!

# two predefined logistic regression models (simple or complex)
# (lrA) simpler logistic regression with 10 terms: stepwise model with . and steps=10
mymodel=glm(Churn~Eqpdays+Retcall+Months+Refurb+Uniqsubs+Mailres+Overage+Mou+Creditde+Actvsubs,data=cell2cell[trainsample,],family='binomial')
# (lrB) more complex logistic regression with 20 terms and interactions: stepwise model with ^2 and steps=20
mymodel=glm(Churn~Eqpdays+Retcall+Months+Refurb+Uniqsubs+Mailres+Overage+Mou+Setprcm+Creditde+Actvsubs+Roam+Changem+Changer+Marryno+Age1+Eqpdays:Months+Months:Mou+Creditde:Changem+Overage:Age1,data=cell2cell[trainsample,],family='binomial')

# give a summary of the model's trained parameters
summary(mymodel)



###############################################################################
### extract information about the model to create new variables
### these new variables are used later in the script
###############################################################################

# create vector of variables used in model called varlist
mvarlist=names(coefficients(mymodel))[-1]   # get the variables used in your logistic regression moodel, except the intercept which is in first position
# check if there are interactions
isinteractions=grepl(":",paste(mvarlist,collapse=" "))
if (isinteractions) {mvarlist=unique(unlist(strsplit(mvarlist,":"))) }
print(mvarlist)  # vector of variables to save

# create matrix of coefficient estimates and std errors and z values
coefdata=summary(mymodel)$coefficients
rownames(coefdata)[1]="Intercept"  # change '(Intercept)' to 'Intercept'
print(coefdata)

# create matrix of data used in the model
modeldata=model.matrix(mymodel,data=cell2cell)
colnames(modeldata)[1]="Intercept"  # change '(Intercept)' to 'Intercept'
mvarcoef=colnames(modeldata)  # list of variables (with potential interactions as ":")
xmodeldata=modeldata  # copy (modeldata has ": and xmodeldata has "X")
if (isinteractions) {
  mvarlisti.idx=grepl(":",mvarcoef)    # get indices of interaction variables in modeldata
  colnames(xmodeldata)=gsub(":","X",colnames(xmodeldata))  # replace : with X for any interactions
}
xmvarcoef=colnames(xmodeldata)
head(xmodeldata)

# create predictions from the logistic regression model
userpred=predict(mymodel,newdata=cell2cell,type='response')  # predict prob for all users
head(userpred)



###############################################################################
### cluster consumers based upon the data weighted by the logistic regression coefficients
### the purpose of this analysis is to identify a smaller set of 'prototypical' customers
###############################################################################

# setup some values
ncluster=20    # number of clusters

# "weight" the data using the logistic regression coefficients.  this allows the cluster to look at
# the variables based upon their contribution to their log-odds ratio
parm=coefdata[,1]  # just extract the parameter estimates in the 1st column
wuserdata=sweep(modeldata,MARGIN=2,parm,"*")  # multiply each row in userdata by parm vector

# let's only consider those users that have a high probability of churn (say the top quartile)
p75=quantile(userpred[!testsample],.75)  # only consider those in the oversampled data
useridx=which(userpred>=p75 & !testsample)  # get the indices of those users that have high churn

# cluster the users into groups
set.seed(612490)   # make sure we get the same solution
grpA=kmeans(wuserdata[useridx,],ncluster,nstart=50)  # add nstart=50 to choose 50 different random seeds

# update the cluster solution with one new group that has everyone else
save.cluster=grpA$cluster   # save the original results
grpA$cluster=rep(ncluster+1,nrow(cell2cell))  # by default place everyone into the last cluster
grpA$cluster[useridx]=save.cluster
grpA$centers=rbind(grpA$centers,colMeans(wuserdata[userpred<p75 & !testsample,]))  # only use the means from oversampled data

# create boxplot of userpred using the clusters
boxplot(userpred[!testsample]~grpA$cluster[!testsample],xlab="Cluster",ylab="Pr(Churn)")
boxplot(cell2cell$Eqpdays[!testsample]~grpA$cluster[!testsample],xlab="Cluster",ylab="Eqpdays")

# create a parallel plot to visualize the centroid values
parallelplot(grpA$centers,auto.key=list(text=as.character(1:nrow(grpA$centers)),space="top",columns=5,lines=T))
parallelplot(grpA$centers,auto.key=list(text=as.character(1:nrow(grpA$centers)),space="top",columns=5,lines=T),common.scale=TRUE)  # choose min and max across variables to give an absolute comparison

# find prototypes of users in each cluster (e.g., observation closest to the centroid)
# compute distances between data and centroids for grpA
# cdist is a matrix between original observation and each of the cluster centroids in the columns
# for example cdist[10,3] would be the distance between observation 10 in cell2cell and cluster #3
# need to use dist package from proxy since it can compute distances between variables
cdistA=proxy::dist(wuserdata,grpA$centers)

# update the clusters for those users in the testsample (by default just assigned to final cluster)
p75.test=quantile(userpred[testsample],.75)
useridx.test=which(userpred>=p75.test & testsample)  # only change those users that are in the top 25%
for (i in useridx.test) {
   grpA$cluster[i]=which.min(cdistA[i,])
}

# find the closest observations to each centroid
#cprototypeB=apply(cdistA,2,which.min)   # this line returns best match but may not be in cluster
cprototypeA=rep(NA,ncol(cdistA))
for (i in 1:ncol(cdistA)) {
  tmpset=which(grpA$cluster==i & !testsample)       # compute subset of customers within segment and in sample
  tmpdist=cdistA[tmpset,i]             # vector of distance values
  tmpmin=min(tmpdist)    # compute the minimum distance within segment
  cprototypeA[i]=tmpset[ intersect(which(tmpdist==tmpmin),seq(tmpdist))[1] ]  # find the index from the tmpset which matches the minimum
}
# print the indices of the observations
print(cprototypeA)
# print the prototypes for each centroid
print(cell2cell[cprototypeA,mvarlist])

# count the number of observations in each centroid
clusterweight=table(grpA$cluster[!testsample])
clusterfrac=clusterweight/sum(clusterweight)  # compute as fraction of total
# compute the weights of each group
clusterchurn=table(grpA$cluster[!testsample],cell2cell$Churn[!testsample])
# adjust the cluster's to represent what would happen in an unadjusted sample
clusterchurn.adj=clusterchurn   # create a copy to store results
clusterchurn.adj[,1]=clusterchurn[,1]*(0.98)/(0.5)  # there are 98% non-churners instead of 50%
clusterchurn.adj[,2]=clusterchurn[,2]*(0.02)/(0.5)  # there are 2% churners instead of 50%
clusterfrac.adj=clusterchurn.adj[,2]/sum(clusterchurn.adj[,2])
# collect summary of appropriate data
clusterresult=cbind(1:length(clusterfrac),clusterfrac,clusterfrac.adj)
colnames(clusterresult)=c("Cluster","PredFrac.unadj","PredFrac")
rownames(clusterresult)=NULL
print(clusterresult)



###############################################################################
### compute adjusted probabilities and LTV in the full dataset
###############################################################################

# set common values for computing ltv
irate=.05/12  # annual discount rate

### compute predictions using your model (for all users)
score=predict(mymodel,newdata=cell2cell)   # compute the score
pchurn.unadj=predict(mymodel,newdata=cell2cell,type='response')
# now adjust predictions to project probability of churn in the original data
adja=pchurn.unadj/(.5/.02)  # the adjustment is the ratio of churn in sample to original
adjb=(1-pchurn.unadj)/(.5/.98)  # the adjustment is the ratio of retain in sample to original
pchurn=adja/(adja+adjb)  # projected prob(churn) in the original dataset
# compute LTV
ltv.unadj=cell2cell$Revenue*(1+irate)/(1+irate-(1-pchurn.unadj))
ltv=cell2cell$Revenue*(1+irate)/(1+irate-(1-pchurn))
#ltv24=(1-((1-pchurn)/(1+irate))^24)*ltv   # what is LTV over 24 months (reduces ltv by amount received after month 24)
promooffer=rep(0,length(pchurn))  # set the promotional offer
promocost=rep(0,length(pchurn))  # set the promotional cost
promotarget=rep(0,length(pchurn))  # set whether to target customer
gain=ltv-promocost  # compute the net gain to ltv after promotional offer

# compare original churn and adjusted (!! if you create more offers include them in result below)
result=cbind(score,pchurn.unadj,adja,adjb,pchurn,ltv,promotarget,promooffer,promocost,gain)
colnames(result)=c("Score","PredChurn.unadj","adja","adjb","PredChurn","LTV","Target","PromoOffer","PromoCost","NetLTV")
head(result)
summary(result)



###############################################################################
### save the results to a file to make it easier to analyze results with Excel
###############################################################################

# create list of users to evaluate
#userlist=1:nrow(cell2cell)   # use this line for full dataset
#userlist=c(119,240,30,32)   # uncomment this list for just our four selected users
#set.seed(123); userlist=sample(1:nrow(cell2cell),100)  # uncomment this line for random sample of 100 users
userlist=cprototypeA  # sets the userlist to prototypes of each of the clusters (e.g., observation nearest centroid)
#userlist=c(1305,259,9520,21351,14001,8265,11129,39767,3441,15170,7661,38200,2085,16975,13821,5329,22013,17067,29304,18851,14007)  # prototypes if using center

# create vector of variables used in model called mvarlist, and add other variables that we want to write out
# these lines require mymodel to be defined above, must either be rpart or glm
evarlist=c("Customer","Revenue","Churn")     # vector of extra variables to save -- regardless of whether they are in the model

# retrieve data about the users  (model.matrix may not work for complex trees)
userdata=cell2cell[userlist,evarlist]  # extract the additional variables that we want
userdata=cbind(UserID=userlist,Cluster=grpA$cluster[userlist],userdata,xmodeldata[userlist,])  # change to dataframe  (can add clusteresult if wanted)
head(userdata)   # print out user data

# combine results with userdata and write to augmented result
augresult=cbind(userdata,result[userlist,],stringsAsFactors=FALSE)  # only selects user in userlist
head(augresult)
augformula=c(paste0(
  "Score=MMULT(",xmvarcoef[1],":",tail(xmvarcoef,1),",coefvec)"),
  "PredChurn.unadj=1/(1+exp(-Score))",
  "adja=PredChurn.unadj/(.5/.02)",
  "adjb=(1-PredChurn.unadj)/(.5/.98)",
  "PredChurn=adja/(adja+adjb)",
  "LTV=Revenue*(1+.05/12)/(1+.05/12-(1-PredChurn))",
  "NetLTV=LTV-PromoCost")
if (isinteractions) {
  augformula=c(augformula,paste(xmvarcoef[mvarlisti.idx],"=",gsub(":","*",mvarcoef[mvarlisti.idx]),sep=""))
}
augresultf=excelformula(augresult,augformula)
head(augresultf)

# create a table with all the data (this is helpful if you want to work with the data in excel)
#write.csv(coefdata,file="cell2cell_coef.csv")       # contains the coefficients
#write.csv(augresult,file="cell2cell_results.csv")   # information about selected users



###############################################################################
### (optional) writes directly to xlsx using openxlsx
### save the results to a file to make it easier to analyze results with Excel
###############################################################################

# create styles
stylegray=createStyle(fgFill="lightgray")
styleyellow=createStyle(fgFill="yellow")
stylegreen=createStyle(fgFill="greenyellow")
styleblue=createStyle(fgFill="lightblue",numFmt="0.000")
styleorange=createStyle(fgFill="orange")
stylered=createStyle(fgFill="lightsalmon")
stylepct=createStyle(numFmt="0.0%")
styledollar=createStyle(numFmt="$ #0.00")
stylefix3=createStyle(numFmt="0.000")
  
# create the workbook
wb=createWorkbook()

# write the sheet with the coefficients
addWorksheet(wb,"Model")
# add format to the data
writeData(wb,sheet=1,x=coefdata,rowNames=TRUE,borders="surrounding")
# name the range of the coefficients (row+1 since there is header, and second column has values)
createNamedRegion(wb,sheet=1,name="coefvec",rows=(1:nrow(coefdata))+1,cols=2)
# add formatting
addStyle(wb,sheet=1,style=stylegray,rows=(1:nrow(coefdata))+1,cols=1,gridExpand=TRUE)
addStyle(wb,sheet=1,style=styleblue,rows=(1:nrow(coefdata))+1,cols=2:5,gridExpand=TRUE)

# add the cluster data to the coefficient spreadsheet
writeData(wb,sheet=1,x=clusterresult,startCol=7,borders="surrounding")
createNamedRegion(wb,sheet=1,name="clusterindex",rows=(1:nrow(clusterresult))+1,cols=7)
createNamedRegion(wb,sheet=1,name="clusterfrac",rows=(1:nrow(clusterresult))+1,cols=9)
addStyle(wb,sheet=1,style=stylegreen,rows=(1:nrow(clusterresult))+1,cols=7:9,gridExpand=TRUE)
addStyle(wb,sheet=1,style=stylepct,rows=(1:nrow(clusterresult))+1,cols=8:9,gridExpand=TRUE,stack=TRUE)

# write the simulator worksheet
addWorksheet(wb,"Simulator")
writeData(wb,sheet=2,x=augresultf,keepNA=TRUE)
# add colors
addStyle(wb,sheet=2,style=stylegray,rows=1,cols=1:ncol(augresultf),gridExpand=TRUE)
addStyle(wb,sheet=2,style=styleorange,rows=(1:nrow(augresultf))+1,cols=match(c("Target","PromoOffer","PromoCost"),colnames(augresultf)),gridExpand=TRUE)
if (isinteractions) {
  addStyle(wb,sheet=2,style=styleyellow,rows=(1:nrow(augresultf))+1,cols=match(xmvarcoef[!mvarlisti.idx],colnames(augresultf)),gridExpand=TRUE)
  addStyle(wb,sheet=2,style=stylered,rows=(1:nrow(augresultf))+1,cols=match(c(xmvarcoef[mvarlisti.idx],"Score","PredChurn.unadj","adja","adjb","PredChurn","LTV","NetLTV"),colnames(augresultf)),gridExpand=TRUE)
} else {
  addStyle(wb,sheet=2,style=styleyellow,rows=(1:nrow(augresultf))+1,cols=match(xmvarcoef,colnames(augresultf)),gridExpand=TRUE)
  addStyle(wb,sheet=2,style=stylered,rows=(1:nrow(augresultf))+1,cols=match(c("Score","PredChurn.unadj","adja","adjb","PredChurn","LTV","NetLTV"),colnames(augresultf)),gridExpand=TRUE)
}
addStyle(wb,sheet=2,style=styledollar,rows=(1:nrow(augresultf))+1,cols=match(c("Revenue","LTV","PromoOffer","PromoCost","NetLTV"),colnames(augresultf)),gridExpand=TRUE,stack=TRUE)
addStyle(wb,sheet=2,style=stylefix3,rows=(1:nrow(augresultf))+1,cols=match(c("Score","adja","adjb"),colnames(augresultf)),gridExpand=TRUE,stack=TRUE)
addStyle(wb,sheet=2,style=stylepct,rows=(1:nrow(augresultf))+1,cols=match(c("PredChurn.unadj","PredChurn"),colnames(augresultf)),gridExpand=TRUE,stack=TRUE)

# write out the user data
alldata=cbind(UserID=1:nrow(modeldata),Cluster=grpA$cluster,cell2cell[,evarlist],modeldata,result[,c("PredChurn","Target","PromoOffer","PromoCost")])  # change to dataframe  (can add clusteresult if wanted)
alldata$Churn[testsample]=NA  # remove the actual churn from the test sample (pretend that we do not know it)
ioffset=2+length(evarlist)

# add user samples
addWorksheet(wb,"Data")
writeData(wb,sheet=3,x=alldata[!testsample,],keepNA=TRUE)
addStyle(wb,sheet=3,style=stylegray,rows=1,cols=1:ncol(alldata),gridExpand=TRUE)
addStyle(wb,sheet=3,style=styleyellow,rows=(1:sum(!testsample))+1,cols=(1:ncol(modeldata))+ioffset,gridExpand=TRUE)
addStyle(wb,sheet=3,style=styleorange,rows=(1:sum(!testsample))+1,cols=match(c("Target","PromoOffer","PromoCost"),colnames(alldata)),gridExpand=TRUE)
addStyle(wb,sheet=3,style=styledollar,rows=(1:sum(!testsample))+1,cols=match(c("Revenue","PromoOffer","PromoCost"),colnames(alldata)),gridExpand=TRUE,stack=TRUE)
addStyle(wb,sheet=3,style=stylepct,rows=(1:sum(!testsample))+1,cols=match(c("PredChurn"),colnames(alldata)),gridExpand=TRUE,stack=TRUE)

# add user samples
addWorksheet(wb,"Pilot")
writeData(wb,sheet=4,x=alldata[testsample,],keepNA=TRUE)
addStyle(wb,sheet=4,style=stylegray,rows=1,cols=1:ncol(alldata),gridExpand=TRUE)
addStyle(wb,sheet=4,style=styleyellow,rows=(1:sum(testsample))+1,cols=(1:ncol(modeldata))+ioffset,gridExpand=TRUE)
addStyle(wb,sheet=4,style=styleorange,rows=(1:sum(testsample))+1,cols=match(c("Target","PromoOffer","PromoCost"),colnames(alldata)),gridExpand=TRUE,stack=TRUE)
addStyle(wb,sheet=4,style=styledollar,rows=(1:sum(testsample))+1,cols=match(c("Revenue","PromoOffer","PromoCost"),colnames(alldata)),gridExpand=TRUE,stack=TRUE)
addStyle(wb,sheet=4,style=stylepct,rows=(1:sum(testsample))+1,cols=match(c("PredChurn"),colnames(alldata)),gridExpand=TRUE,stack=TRUE)

# write workbook now
saveWorkbook(wb, file="cell2cell_logistic_20.xlsx", overwrite=TRUE)

