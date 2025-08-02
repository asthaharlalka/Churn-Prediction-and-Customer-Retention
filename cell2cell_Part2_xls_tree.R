###############################################################################
# Script: cell2cell_Part2_xls_tree.R
# Copyright (c) 2019 by Alan Montgomery. Distributed using license CC BY-NC 4.0
# To view this license see https://creativecommons.org/licenses/by-nc/4.0/
#
# creates a spreadsheet "cell2cell_results.csv" using your 'best' tree model
# from Part1 to predict the churn and LTV, which can be used to predict how
# consumers will respond to potential offers or scenarios
# (e.g., give everyone a free phone)
#
# also to help reduce the complexity of the problem it outputs results only
# for a selected set of customers (instead of all 40,000)
# The choice of customers is to find the customer that is closest to the 
# centroid from each branch of the tree
###############################################################################


###############################################################################
### setup
###############################################################################

# setup environment, make sure this library has been installed
if (!require(rstudioapi)) {install.packages("rstudioapi"); library(rstudioapi)}
if (!require(tree)) {install.packages("tree"); library(tree)}
if (!require(rpart)) {install.packages("rpart"); library(rpart)}
if (!require(rpart.plot)) {install.packages("rpart.plot"); library(rpart.plot)}
if (!require(rpart.utils)) {install.packages("rpart.utils"); library(rpart.utils)}
if (!require(plyr)) {install.packages("plyr"); library(plyr)}
if (!require(stringr)) {install.packages("stringr"); library(stringr)}
if (!require(psych)) {install.packages("psych"); library(psych)}
if (!require(proxy)) {install.packages("proxy"); library(proxy)}
if (!require(openxlsx)) {install.packages("openxlsx"); library(openxlsx)}

# function to predict nodes for a new dataset with rpart
# https://stackoverflow.com/questions/29304349/how-to-get-terminal-nodes-for-a-new-observation-from-an-rpart-object
predict_nodes <-
  function (object, newdata, na.action = na.pass) {
    where <-
      if (missing(newdata)) 
        object$where
    else {
      if (is.null(attr(newdata, "terms"))) {
        Terms <- delete.response(object$terms)
        newdata <- model.frame(Terms, newdata, na.action = na.action, 
                               xlev = attr(object, "xlevels"))
        if (!is.null(cl <- attr(Terms, "dataClasses"))) 
          .checkMFClasses(cl, newdata, TRUE)
      }
      rpart:::pred.rpart(object, rpart:::rpart.matrix(newdata))
    }
    as.integer(row.names(object$frame))[where]
  }

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
### estimate your tree model
### only estimates the model, does not do any diagnostics or statistics
### you must save your model to "mymodel" for the rest of the script
### !! use your best model from part 1
###############################################################################

# !! just uncomment one model OR paste your "mymodel=" below and comment out the others !!

# two predefined tree models (simple or complex)
ctree.full = rpart(Churn~., data=cell2cell[trainsample,], control=rpart.control(cp=0.0005),model=TRUE)
ctreeA=prune(ctree.full,cp=0.005)  # prune tree using chosen complexity parameter !! simple tree (treeA) !!
ctreeC=prune(ctree.full,cp=0.002)  # prune tree using chosen complexity parameter
ctreeB=prune(ctree.full,cp=0.00090890)  # prune tree using chosen complexity parameter !! better tree (treeB) !!
ctreeD=prune(ctree.full,cp=0.0012)  # prune tree using chosen complexity parameter !! better tree (treeB) !!
mymodel = ctreeD

# give summary of the model
summary(mymodel)

# save the prp plot to a jpeg file
jpeg(filename="cell2cell_tree.jpeg",width=2048,height=1536,units="px",res=288)
prp(mymodel,extra=101,nn=TRUE,xcompact=FALSE,tweak=1.1)  # add the size and proportion of data in the node
dev.off()



###############################################################################
### extract information about the model to create new variables
### these new variables are used later in the script
###############################################################################

# create vector of variables used in model called varlist
#mvarlist=names(mymodel$variable.importance)   # get the list of variables from rpart by looking at the importance
mvarlist=unique(as.character(mymodel$frame$var))  # detect only variables that occur
mvarlist=mvarlist[mvarlist!="<leaf>"]  # remove the <leaf> label
print(mvarlist)  # vector of variables to save

# create matrix of data used in the model
modeldata=as.data.frame(cell2cell[,mvarlist])      # extract the variables directly from the data
head(modeldata)

# create predictions from the logistic regression model
userpred=predict(mymodel,newdata=cell2cell,type='vector')  # predict prob for all users
head(userpred)

# save the actual churn labels
truechurn = cell2cell$Churn

# extract the nodes associated with the trees
prnode=predict_nodes(mymodel,cell2cell)    # uses the predict_nodes function, returns vector of leafs
namenodes=unique(prnode)     # get the names of the nodes
namenodes=namenodes[order(namenodes)]   # sort the nodes in increasing order
prnodeidx=match(prnode,namenodes)  # get the corresponding index from prnode in namenodes
numnodes=length(namenodes)   # count the number of terminal nodes (like the # of clusters)



###############################################################################
### find a prototypical user in each tree branch, two variables are set:
###   centers = matrix of means (nrow=numnodes X ncol=length(mvarlist))
###   prototypeTree = vector of closest observation to each center (length=numnodes)
###############################################################################

# compute the centroids for all observations in a leaf (e.g., cluster)
centers=aggregate(cell2cell[!testsample,mvarlist],list(prnode[!testsample]),mean)
# replace the mean with the mode for any factors
mvarlist.isfactor=sapply(cell2cell[1,mvarlist],is.factor)
for (i in names(mvarlist.isfactor)[mvarlist.isfactor]) {
  tablemode=function(x) {names(table(x)[which.max(table(x))])}  # return a string with the modal value of an input factor
  centers[,i]=factor(aggregate(cell2cell[!testsample,i],list(prnode[!testsample]),tablemode)[,2],levels=levels(cell2cell[!testsample,i]))  # return factor using original levels
}

# for each leaf (final nodes in the tree) let's find the person that is closest to the center
prototypeTree=rep(NA,numnodes)  # create a place to store the nearest neighbor for each node
for (i in 1:numnodes) {
  # indexes of obserations in node
  nodeidx=which(prnode==namenodes[i] & !testsample)
  # compute distances to center of each cluster/node
  cdistnode=proxy::dist(cell2cell[nodeidx,mvarlist],centers[i,mvarlist])
  # find the closest observations to each centroid
  prototypeTree[i]=nodeidx[which.min(cdistnode)]
}



###############################################################################
### export the decision tree rules to excel
###############################################################################

# extract the notes associated with the trees
leaf.pred=prnode[trainsample]  # use the training sample since we want estimated predictions to match with ctree
leaf.total=table(leaf.pred)  # compute total number of observations in leaf
result=aggregate(truechurn[trainsample],by=list(leaf=leaf.pred),mean)  # compute average pr(churn) by leaf
leaf.name=as.character(result$leaf)  # get the names
leaf.probchurn=result$x; names(leaf.probchurn)=leaf.name
leaf.probchurn.adj=(leaf.probchurn/(.5/.02))/(leaf.probchurn/(.5/.02)+(1-leaf.probchurn)/(.5/.98))  # adjust for oversampling
leaf.frac=prop.table(table(leaf.pred))  # compute the relevance of each node (e.g. proportion of sample)
leaf.predchurn=leaf.probchurn[as.character(leaf.pred)]  # lookup the predictions
leaf.frac.adj=(leaf.probchurn*.02+(1-leaf.probchurn)*.98)*leaf.frac/sum((leaf.probchurn*.02+(1-leaf.probchurn)*.98)*leaf.frac)  # reweight by the adjusted sample sizes

# extract the rules
leaf.rules=rpart.rules(mymodel)  # returns list that gives rules to evaluate for final leaf
leaf.rulesdf=data.frame(leaf=leaf.name,rules=unlist(leaf.rules[as.numeric(leaf.name)]),stringsAsFactors=FALSE)  # list rules in wide form
leaf.cond=rpart.subrules.table(mymodel)  # subrules in long form

# create excel formula for each subrule
leaf.condxls=data.frame(Subrule=as.character(unique(leaf.cond$Subrule)),formula="",stringsAsFactors=FALSE)
for (i in 1:nrow(leaf.condxls) ) {
  # extract subrules
  isubrule=subset(leaf.cond,Subrule %in% leaf.condxls$Subrule[i])
  # create an excel formula based upon the subrule
  if (nrow(isubrule)>1) {
    ivariable=as.character(isubrule$Variable)[1]
    mstring=paste0(paste0('"',isubrule$Value,'"'),collapse=',')
    iformula=paste0("=ISNUMBER(MATCH(",ivariable,",{",mstring,"},0))")
  } else if (!is.na(isubrule$Greater)) {
    ivariable=as.character(isubrule$Variable)
    iformula=paste0("=IF(",ivariable,">=",isubrule$Greater,",TRUE,FALSE)")
  } else if (!is.na(isubrule$Less)) {
    ivariable=as.character(isubrule$Variable)
    iformula=paste0("=IF(",ivariable,"<",isubrule$Less,",TRUE,FALSE)")
  }
  # update the ouput now
  leaf.condxls[i,"variable"]=ivariable
  leaf.condxls[i,"formula"]=iformula
}
leaf.condxls$Subrule=paste0("COND",as.character(leaf.condxls$Subrule))  # prefix all conditions with COND
rownames(leaf.condxls)=leaf.condxls$Subrule

# create excel formula for each rule
leaf.rulesxls=data.frame(leaf=leaf.rulesdf$leaf,probchurn.unadj=as.numeric(leaf.probchurn),frac.unadj=as.numeric(leaf.frac),
                         probchurn=as.numeric(leaf.probchurn.adj),frac=as.numeric(leaf.frac.adj),formula="",
                         row.names=leaf.rulesdf$leaf,stringsAsFactors=FALSE)
for (i in 1:nrow(leaf.rulesxls) ) {
  # create an excel formula based upon the subrule
  mstring=paste0( paste0("COND",strsplit(leaf.rulesdf$rules[i],',')[[1]]), collapse=',')
  iformula=paste0("=AND(",mstring,")")
  # update the ouput now
  leaf.rulesxls[i,"formula"]=iformula
}
rownames(leaf.rulesxls)=paste0("RULE",leaf.rulesxls$leaf)
colnames(leaf.rulesxls)=c("LeafRule","PredChurn.unadj","PredFrac.unadj","PredChurn","PredFrac","formula")



###############################################################################
### compute adjusted probabilities and LTV in the full dataset
###
### hint:
### you may want to create new campaigns as below for offerA
### you can then decide which offer to select for each user by looking to
### see which offer gives you the maximum return
###############################################################################

# set common values for computing ltv
irate=.05/12  # annual discount rate

### compute predictions using your model (for all users)
pchurn.unadj=predict(mymodel,newdata=cell2cell,type='vector')
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
result=cbind(prnodeidx,prnode,pchurn.unadj,adja,adjb,pchurn,ltv,promotarget,promooffer,promocost,gain)
colnames(result)=c("LeafIndex","LeafRule","PredChurn.unadj","adja","adjb","PredChurn","LTV","Target","PromoOffer","PromoCost","NetLTV")
head(result)
summary(result)



###############################################################################
### save the results to a file to make it easier to analyze results with Excel
###############################################################################

# create list of users to evaluate
#userlist=1:nrow(cell2cell)   # use this line for full dataset
#userlist=c(119,240,30,32)   # uncomment this list for just our four selected users
#set.seed(123); userlist=sample(1:nrow(cell2cell),100)  # uncomment this line for random sample of 100 users
userlist=prototypeTree  # sets the userlist to prototypes of each of the clusters (e.g., observation nearest centroid)
#userlist=c(1305,259,9520,21351,14001,8265,11129,39767,3441,15170,7661,38200,2085,16975,13821,5329,22013,17067,29304,18851,14007)  # prototypes for default logistic clusters
#userlist=c(25080,30046,21396,20143,36988,1547,14690,34360,13197,5665,25120,14479,8411,30705,17980,32302,34272)  # prototypes for default tree

# create vector of variables used in model called mvarlist, and add other variables that we want to write out
# these lines require mymodel to be defined above, must either be rpart or glm
evarlist=c("Customer","Revenue","Churn")     # vector of extra variables to save -- regardless of whether they are in the model

# retrieve data about the users  (model.matrix may not work for complex trees)
userdata=cell2cell[userlist,evarlist]  # extract the additional variables that we want
userdata=cbind(UserID=userlist,userdata,modeldata[userlist,])  # change to dataframe (can add clusterresult if desired)
print(userdata)   # print out user data
head(userdata)

# create formulas for conditions and rules (repeat the conditions and rules, will substitute cell values below)
conddata=matrix(leaf.condxls$formula,nrow=numnodes,ncol=nrow(leaf.condxls),dimnames=list(userlist,rownames(leaf.condxls)),byrow=TRUE)
ruledata=matrix(leaf.rulesxls$formula,nrow=numnodes,ncol=nrow(leaf.rulesxls),dimnames=list(userlist,rownames(leaf.rulesxls)),byrow=TRUE)

# combine results with userdata and write to augmented result (creates a data frame since userdata is df)
augresult=cbind(userdata,conddata,ruledata,result[userlist,],stringsAsFactors=FALSE)  # only selects user in userlist
head(augresult)
augxname=excelcol(colnames(augresult))  # create variable with appropriate excel column names for augresult

# create an excel formula version of result (this is a vector not a dataframe)
resultformula=c(paste0("=MATCH(TRUE,",paste(colnames(ruledata)[c(1,ncol(ruledata))],collapse=":"),",0)"),
                "=INDEX(LeafVec,LeafIndex)","=INDEX(UnadjProbVec,LeafIndex)",
                "=PredChurn.unadj/(.5/.02)","=(1-PredChurn.unadj)/(.5/.98)","=adja/(adja+adjb)",
                "=Revenue*(1+.05/12)/(1+.05/12-(1-PredChurn))","0","0","0","=LTV-PromoCost")
names(resultformula)=colnames(result)
resultfmatrix=matrix(resultformula,nrow=nrow(augresult),ncol=length(resultformula),byrow=TRUE,dimnames=list(NULL,names(resultformula)))

# replace the variable names in conditions with appropriate column name
augformula=cbind(conddata,ruledata,resultfmatrix)  # create copy of formulas to change
# replace variable names from conditions and rules with appropriate Excel variable
for (i in 1:nrow(augresult)) {
  # append condcol with row number+1 (needs to be incremented to adjust for row with names)
  augxrow=paste0(augxname,i+1)
  names(augxrow)=paste0("\\b",names(augxname),"\\b")  # use \\b to only replace words
  # need to replace variable names in formulas with appropriate excel column
  augformula[i,]=str_replace_all(augformula[i,],augxrow)
}

# combine results with userdata and write to augmented result (creates a data frame since userdata is df)
augresult=cbind(userdata,augformula,stringsAsFactors=FALSE)  # only selects user in userlist
head(augresult)

# create a table with all the data (this is helpful if you want to work with the data in excel)
#write.csv(augresult,file="cell2cell_tree_results.csv")   # information about selected users
#write.csv(leaf.rulesxls,file="cell2cell_tree_rules.csv")
#write.csv(leaf.condxls,file="cell2cell_tree_cond.csv")



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

# change the data frame so these columns are identified as formula
class(leaf.condxls$formula)=c(class(leaf.condxls$formula),"formula")
class(leaf.rulesxls$formula)=c(class(leaf.rulesxls$formula),"formula")

# create a template for the data
datatemp=t(cell2cell[1,mvarlist])
colnames(datatemp)="Values"

# create a template for predictions
predtemp=data.frame(
  formula=c("MATCH(TRUE,RulesVec,0)","INDEX(LeafVec,LeafIndex)","INDEX(UnadjProbVec,LeafIndex)","INDEX(AdjProbVec,LeafIndex)"),
  row.names=c("LeafIndex","LeafRule","PredChurn.Unadj","PredChurn"))
class(predtemp$formula)=c(class(predtemp$formula),"formula")

# write the sheet with the rules
addWorksheet(wb,"Model")
writeData(wb,"Model",x=datatemp,rowNames=TRUE,borders="surrounding")
writeData(wb,"Model",x=leaf.condxls[,c("Subrule","formula")],startCol=4,borders="surrounding")
writeData(wb,"Model",x=leaf.rulesxls,startCol=7,borders="surrounding")
writeData(wb,"Model",x=predtemp,rowNames=TRUE,startCol=7,startRow=nrow(leaf.rulesxls)+3,borders="surrounding")
insertImage(wb, "Model", "cell2cell_tree.jpeg", width=11.18, height=7.82, units="in",startCol=1,startRow=nrow(leaf.rulesxls)+9)
# name the data
for (i in 1:nrow(datatemp)) {
  createNamedRegion(wb,sheet=1,name=mvarlist[i],rows=i+1,cols=2)
}
# name each of the conditions (row+1 since there is header, and second column has values)
for (i in 1:nrow(leaf.condxls)) {
  createNamedRegion(wb,sheet=1,name=leaf.condxls[i,1],rows=i+1,cols=5)
}
# name the rules
createNamedRegion(wb,sheet=1,name="LeafVec",rows=(1:nrow(leaf.rulesxls))+1,cols=7)
createNamedRegion(wb,sheet=1,name="UnadjProbVec",rows=(1:nrow(leaf.rulesxls))+1,cols=8)
createNamedRegion(wb,sheet=1,name="UnadjFracVec",rows=(1:nrow(leaf.rulesxls))+1,cols=9)
createNamedRegion(wb,sheet=1,name="AdjProbVec",rows=(1:nrow(leaf.rulesxls))+1,cols=10)
createNamedRegion(wb,sheet=1,name="AdjFracVec",rows=(1:nrow(leaf.rulesxls))+1,cols=11)
createNamedRegion(wb,sheet=1,name="RulesVec",rows=(1:nrow(leaf.rulesxls))+1,cols=12)
# name each of the predictions
for (i in 1:nrow(predtemp)) {
  createNamedRegion(wb,sheet=1,name=rownames(predtemp)[i],rows=i+nrow(leaf.rulesxls)+3,cols=8)
}
# add colours
addStyle(wb,sheet=1,style=stylegray,rows=(1:length(datatemp))+1,cols=1,gridExpand=TRUE)
addStyle(wb,sheet=1,style=styleyellow,rows=(1:length(datatemp))+1,cols=2,gridExpand=TRUE)
addStyle(wb,sheet=1,style=stylegray,rows=(1:nrow(leaf.condxls))+1,cols=4,gridExpand=TRUE)
addStyle(wb,sheet=1,style=styleblue,rows=(1:nrow(leaf.condxls))+1,cols=5,gridExpand=TRUE)
addStyle(wb,sheet=1,style=stylegreen,rows=(1:nrow(leaf.rulesxls))+1,cols=7:11,gridExpand=TRUE)
addStyle(wb,sheet=1,style=styleblue,rows=(1:nrow(leaf.rulesxls))+1,cols=12,gridExpand=TRUE)
addStyle(wb,sheet=1,style=stylegray,rows=(nrow(leaf.rulesxls)+4):(nrow(leaf.rulesxls)+3+nrow(predtemp)),cols=7,gridExpand=TRUE)
addStyle(wb,sheet=1,style=stylered,rows=(nrow(leaf.rulesxls)+4):(nrow(leaf.rulesxls)+3+nrow(predtemp)),cols=8,gridExpand=TRUE)
addStyle(wb,sheet=1,style=stylepct,rows=(1:nrow(leaf.rulesxls))+1,cols=8:11,gridExpand=TRUE,stack=TRUE)

# convert the conditions and rules into formula
augxname=excelcol(colnames(augresult))
for (i in c(rownames(leaf.condxls),rownames(leaf.rulesxls),colnames(result))) {
  class(augresult[,i])=c(class(augresult[,i]),"formula")
}

# write out the worksheet now
addWorksheet(wb,"Simulator")
writeData(wb,sheet=2,x=augresult)
# add colours
augrows=(1:nrow(augresult))+1
ioffset1=ncol(userdata)
ioffset2=ioffset1+ncol(conddata)+ncol(ruledata)
addStyle(wb,sheet=2,style=stylegray,rows=1,cols=1:ncol(augresult),gridExpand=TRUE)
addStyle(wb,sheet=2,style=styleyellow,rows=augrows,cols=5:ncol(userdata),gridExpand=TRUE)
addStyle(wb,sheet=2,style=styleblue,rows=augrows,cols=(ioffset1+1):(ioffset1+ncol(conddata)+ncol(ruledata)),gridExpand=TRUE)
addStyle(wb,sheet=2,style=stylered,rows=augrows,cols=match(c("LeafIndex","LeafRule","PredChurn.unadj","adja","adjb","PredChurn","LTV","NetLTV"),colnames(augresult)),gridExpand=TRUE)
addStyle(wb,sheet=2,style=styleorange,rows=augrows,cols=match(c("Target","PromoOffer","PromoCost"),colnames(augresult)),gridExpand=TRUE)
addStyle(wb,sheet=2,style=styledollar,rows=augrows,cols=match(c("Revenue","LTV","PromoOffer","PromoCost","NetLTV"),colnames(augresult)),gridExpand=TRUE,stack=TRUE)
addStyle(wb,sheet=2,style=stylefix3,rows=augrows,cols=match(c("Score","adja","adjb"),colnames(augresult)),gridExpand=TRUE,stack=TRUE)
addStyle(wb,sheet=2,style=stylepct,rows=augrows,cols=match(c("PredChurn.unadj","PredChurn"),colnames(augresult)),gridExpand=TRUE,stack=TRUE)

# write out the user data
alldata=cbind(UserID=1:nrow(modeldata),LeafIndex=prnodeidx,LeafRule=prnode,cell2cell[,evarlist],modeldata,result[,c("PredChurn","Target","PromoOffer","PromoCost")])  # change to dataframe  (can add clusteresult if wanted)
alldata$Churn[testsample]=NA  # remove the actual churn from the test sample (pretend that we do not know it)
ioffset=3+length(evarlist)

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
saveWorkbook(wb, file="cell2cell_tree_cp_0012.xlsx", overwrite=TRUE)

