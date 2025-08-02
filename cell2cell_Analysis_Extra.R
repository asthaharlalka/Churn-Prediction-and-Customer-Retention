###############################################################################
# Script: cell2cell_Analysis_Extra.R
# Copyright (c) 2020 by Alan Montgomery. Distributed using license CC BY-NC 4.0
# To view this license see https://creativecommons.org/licenses/by-nc/4.0/
#
### Script to to optional activities.  This is not required.  Below is the
### segments of code.  This code assumes you have run Part1 of the script.
###
### @moreplots: alternative plots for describing the data
### @moretreeplot: plotting very small trees using manual pruning and fancier color tree plots
### @plotlr: plotting logistic regression
### @clusterlr: 
###    cluster consumers based upon the data weighted by the logistic regression coefficients
###    the purpose of this analysis is to identify a smaller set of 'prototypical' customers
###    assumes that lrmdl has been estimated already (see Part1.R)
### @whylrchurn: look for the three most likely reasons customer will churn
### @prototypetree: find a prototypical user in each tree branch
### @randomforest: estimate a random forest
### @boostedtree: estimate a boosted tree with gbm
### @graph models: overlay the predictions from the models against the actual data
### @compare: compare models using ROC plot
### @whatif
###    evaluate predictions for selected users for a variety of offers
###    this section of code creates a number of alternative offers by changing
###    the independent variables and then compute pr(churn) and LTV
###############################################################################



###############################################################################
### setup
###############################################################################

# setup environment, make sure this library has been installed
if (!require(tree)) {install.packages("tree"); library(tree)}
# setup environment (if you want to use fancy tree plots)
if (!require(rpart)) {install.packages("rpart"); library(rpart)}
if (!require(rattle)) {install.packages("rattle"); library(rattle)}
if (!require(rpart.plot)) {install.packages("rpart.plot"); library(rpart.plot)}
if (!require(RColorBrewer)) {install.packages("RColorBrewer"); library(RColorBrewer)}
if (!require(party)) {install.packages("party"); library(party)}
if (!require(partykit)) {install.packages("partykit"); library(partykit)}
# a better scatterplot matrix routine
if (!require(car)) {install.packages("car"); library(car)}
# better summary tables
if (!require(psych)) {install.packages("psych"); library(psych)}
# tools for random forest
if (!require(randomForest)) {install.packages("randomForest"); library(randomForest)} 
# tools for data editing
if (!require(plyr)) {install.packages("plyr"); library(plyr)} 
# tools for logistic regression
if (!require(visreg)) {install.packages("visreg"); library(visreg)}  # visualize regression
if (!require(ROCR)) {install.packages("ROCR"); library(ROCR)}  # ROC curve for tree or logistic
if (!require(plotmo)) {install.packages("plotmo"); library(plotmo)}
# setup environment, for plots
if (!require(reshape2)) {install.packages("reshape2"); library(reshape2)}
if (!require(gplots)) {install.packages("gplots"); library(gplots)}
if (!require(ggplot2)) {install.packages("ggplot2"); library(ggplot2)}
# data manipulation
if (!require(plyr)) {install.packages("plyr"); library(plyr)}

# define a function to summary a classification matrix we will use later
# this is a function which has three arguments:
#    predprob  = vector of prediction probabilities
#    predclass = vector of predicted labels
#    trueclass = vector of true class labels
confmatrix.summary <- function(predprob,predclass,trueclass) {
  # compute confusion matrix (columns have truth, rows have predictions)
  results = xtabs(~predclass+trueclass)  
  # compute usual metrics from the confusion matrix
  accuracy = (results[1,1]+results[2,2])/sum(results)   # how many correct guesses along the diagonal
  truepos = results[2,2]/(results[1,2]+results[2,2])  # how many correct "churn" guesses
  precision = results[2,2]/(results[2,1]+results[2,2]) # proportion of correct positive guesses 
  trueneg = results[1,1]/(results[2,1]+results[1,1])  # how many correct "non-churn" guesses 
  # compute the lift using the predictions for the 10% of most likely
  topchurn = as.vector( predprob >= as.numeric(quantile(predprob,probs=.9)))  # which customers are most likely to churn
  ( baseconv=sum(trueclass==1)/length(trueclass) )  # what proportion would we have expected purely due to chance
  ( actconv=sum(trueclass[topchurn])/sum(topchurn))  # what proportion did we actually predict
  ( lift=actconv/baseconv )  # what is the ratio of how many we got to what we expected
  return(list(confmatrix=results,accuracy=accuracy,truepos=truepos,precision=precision,trueneg=trueneg,lift=lift))
}


###############################################################################
### read in the data and prepare the dataset for analysis
###############################################################################

# set to working directory of script (assumes data in same directory as script)
#setwd(dirname(rstudioapi::getActiveDocumentContext()$path))  # only works in Rstudio scripts
# import dataset from file (change the directory to where your data is stored)
setwd("~/Documents/class/analytical marketing/cases/cell2cell/data")  # !! change this to your directory !!

# import dataset from file
cell2cell=read.csv("cell2cell_data.csv")
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
plotsample=sample(1:nrow(cell2cell),200)

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

# choose some common variables
svarlist=c("Churn","Eqpdays","Months","Recchrge","Revenue","Csa","Customer","Age1","Age2","Mailflag","Retcall")
nvarlist=c("Churn","Eqpdays","Months","Revenue","Customer","Age1")  # numeric variables only



###############################################################################
### estimate a tree model with all variables
###############################################################################

# estimate a model with all the variables that is very deep
ctree.full = rpart(Churn~., data=cell2cell[trainsample,], control=rpart.control(cp=0.0005), model=TRUE)
summary(ctree.full)
# uncomment the line below to view the full tree -- clearly needs pruning -- which is what the commands below do
#prp(ctree.full)  # make sure your plot window is large or this command can cause problems

# these lines are helpful to find the "best" value of cp
# A good choice of cp for pruning is often the leftmost value for which the mean lies below the horizontal line.
printcp(ctree.full)               # display table of optimal prunings based on complexity parameter
plotcp(ctree.full)                # visualize cross-validation results

# prune the tree back !! choose on of the lines below for treeA or treeB, and leave the other commented out !!
#ctree=prune(ctree.full,cp=0.005)  # prune tree using chosen complexity parameter !! simple tree (treeA) !!
ctree=prune(ctree.full,cp=0.00090890)  # prune tree using chosen complexity parameter !! better tree (treeB) !!

# visualize the trees 
par(mfrow=c(1,1))         # reset one graphic per panel
#plot(ctree); text(ctree)  # simple graph
prp(ctree,extra=101,nn=TRUE)  # add the size and proportion of data in the node
#fancyRpartPlot(ctree)     # fancy graphic  !! uncomment if you load library(rattle) !!
plotmo(ctree)             # evaluates selected input but holds other values at median
#plotmo(ctree,pmethod="apartdep")   # evaluates selected input and averages other values  !! pmethod="partdep" is better but slower !!

# compute predictions for the entire sample -- notice we only train on trainsample
pchurn.tree = predict(ctree,newdata=cell2cell,type='vector')
cchurn.tree = (pchurn.tree>.5)+0   # make our predictions using a 50% cutoff, !! this can threshold can be changed !!
truechurn = cell2cell$Churn

# compute confusion matrix and some usual statistics (uncomment train and prediction samples if you want to see these statistics)
#confmatrix.summary(pchurn.tree[trainsample],cchurn.tree[trainsample],truechurn[trainsample])  # summary for training sample (look to see if train is substantially better than valid)
confmatrix.summary(pchurn.tree[validsample],cchurn.tree[validsample],truechurn[validsample])  # summary for validation sample
#confmatrix.summary(pchurn.tree[predsample],cchurn.tree[predsample],truechurn[predsample]) # summary for prediction sample (use this when you want to know how good your "final" model is)

# compute ROC and AUC
rocpred.tree = prediction(pchurn.tree[validsample],cell2cell$Churn[validsample])  # compute predictions using "prediction"
rocperf.tree = performance(rocpred.tree, measure = "tpr", x.measure = "fpr")
plot(rocperf.tree, col=rainbow(10)); abline(a=0, b= 1)
auc.tmp = performance(rocpred.tree,"auc")  # compute area under curve
(auc.tree = as.numeric(auc.tmp@y.values))



###############################################################################
### estimate a logistic regression model
### you can either use a prespecified model,
### or use stepwise regression model with all the variables and their interactions
###############################################################################

# !! either use the first block of code to create your own regression,  !!
# !! or uncomment the lrmdl in the second block of code to use a prespecified model !!

# run a step-wise regression
# first estimate the null model (this just has an intercept)
##null = glm(Churn~1,data=cell2cell[trainsample,],family='binomial')
# second estimate a complete model (with all variables that you are interested in)
##full = glm(Churn~.,data=cell2cell[trainsample,],family='binomial')  # takes a long time
# if you have time uncomment the following line and include all squared terms (e.g., nonlinear effects)
##full = glm(Churn~.^2,data=cell2cell[plotsample,],family='binomial')  # takes a very long time -- but since we just want the formula for stepwise can just use plotsample instead of trainsample
# finally estimate the step wise regression starting with the null model
##lrmdl = step(null, scope=formula(full),steps=15,dir="forward")  # !! can increase beyond 15 steps, just takes more time

# this logistic regression uses some important variables and is a gives a good model
# if you uncomment the stepwise regression then comment out the following line
# (example) simpler logistic regression used in logistic regression simulator
#lrmdl=glm(Churn~Eqpdays+Retcall+Months+Overage+Mou+Changem,data=cell2cell[trainsample,],family='binomial')
# (simple) simpler logistic regression with 10 terms: stepwise model with . and steps=10
#lrmdl=glm(Churn~Eqpdays+Retcall+Months+Refurb+Uniqsubs+Mailres+Overage+Mou+Creditde+Actvsubs,data=cell2cell[trainsample,],family='binomial')
# (base: lrB) logistic regression with 20 terms and interactions: stepwise model with ^2 and steps=20, note: Eqpdays:Months means Eqpdays*Months in the model
lrmdl=glm(Churn~Eqpdays+Retcall+Months+Refurb+Uniqsubs+Mailres+Overage+Mou+Setprcm+Creditde+Actvsubs+Roam+Changem+Changer+Marryno+Age1+Eqpdays:Months+Months:Mou+Creditde:Changem+Overage:Age1,data=cell2cell[trainsample,],family='binomial')

# give a summary of the model's trained parameters
summary(lrmdl)
plotmo(lrmdl)             # evaluates selected input but holds other values at median
#plotmo(lrmdl,pmethod="partdep")   # evaluates selected input and averages other values  !! pmethod="apartdep" is faster but approximate !!

# compute predictions for the entire sample -- but model was only trained on trainsample
pchurn.lr = predict(lrmdl,newdata=cell2cell,type='response')
cchurn.lr = (pchurn.lr>.5)+0
truechurn = cell2cell$Churn

# compute confusion matrix and some usual statistics (uncomment train and prediction samples if you want to see these statistics)
#confmatrix.summary(pchurn.lr[trainsample],cchurn.lr[trainsample],truechurn[trainsample])  # summary for training sample (look to see if train is substantially better than valid)
confmatrix.summary(pchurn.lr[validsample],cchurn.lr[validsample],truechurn[validsample])  # summary for validation sample
#confmatrix.summary(pchurn.lr[predsample],cchurn.lr[predsample],truechurn[predsample]) # summary for prediction sample (use this when you want to know how good your "final" model is)

# compute ROC and AUC
rocpred.lr = prediction(pchurn.lr[validsample],cell2cell$Churn[validsample])  # compute predictions using "prediction"
rocperf.lr = performance(rocpred.lr, measure = "tpr", x.measure = "fpr")
plot(rocperf.lr, col=rainbow(10)); abline(a=0, b= 1)
auc.tmp = performance(rocpred.lr,"auc")  # compute area under curve
(auc.lr = as.numeric(auc.tmp@y.values))



#@moreplots
###############################################################################
### alternative plots for describing the data
###############################################################################

# boxplots on a log scale makes it easier to see relative differences
par(mfrow=c(2,4),mar=c(5,5,1,1))
boxplot(Eqpdays~Churn,data=cell2cell[plotsample,],xlab="Churn",ylab="Eqpdays",log='y')
boxplot(Months~Churn,data=cell2cell[plotsample,],xlab="Churn",ylab="Months",log='y')
boxplot(Recchrge~Churn,data=cell2cell[plotsample,],xlab="Churn",ylab="Recchrge",log='y')
boxplot(Revenue~Churn,data=cell2cell[plotsample,],xlab="Churn",ylab="Revenue",log='y')
boxplot(Customer~Churn,data=cell2cell[plotsample,],xlab="Churn",ylab="Customer",log='y')
boxplot(Age1~Churn,data=cell2cell[plotsample,],xlab="Churn",ylab="Age1",log='y')
boxplot(Age2~Churn,data=cell2cell[plotsample,],xlab="Churn",ylab="Age2",log='y')
par(mfrow=c(1,1))



#@moretreeplot
###############################################################################
### plotting very small trees using manual pruning and fancier color tree plots
###############################################################################

# manual method for pruning the tree, mainly helpful if you have a really big tree and want to better plot part of it
# automatic prune should be good, but another use when you have big trees is to 
# clip branches of the tree so you can better see branches (in plot window clip on node(s) and then finish)
manual.prune=prp(ctree,snip=TRUE)$obj  # if you want to manually snip the tree click on node to delete
prp(manual.prune,extra=101)    # plot your snipped tree
#fancyRpartPlot(manual.prune)  # plot your snipped tree (!! uncomment for a fancy plot !!)
#ctree=manual.prune  # uncomment to save your manually pruned tree (careful this overwrites your best ctree from above)

# tree plot that colors in branches by probability
# source is Milborrow (2016) at http://www.milbo.org/doc/prp.pdf
heat.tree <- function(tree, low.is.green=FALSE, ...) { # dots args passed to prp
  y <- tree$frame$yval
  if(low.is.green)
    y <- -y
  max <- max(y)
  min <- min(y)
  cols <- rainbow(99, end=.36)[
    ifelse(y >  y[1], (y-y[1]) * (99-50) / (max-y[1]) + 50,
           (y-min)  * (50-1)  / (y[1]-min) + 1)]
  prp(tree, branch.col=cols, box.col=cols, ...)
}
heat.tree(ctree, type=4, varlen=0, faclen=0, fallen.leaves=TRUE, cex=.8)
heat.tree(ctree, varlen=0, faclen=0, fallen.leaves=TRUE, branch.type=5, yesno=FALSE, cex=.7)  # width proportional to # of observations



#@plotlr
###############################################################################
### plotting logistic regression
###############################################################################

# plot the log of the odds ratio of churn as function of equipment days
par(mfrow=c(2,1))
visreg(lrmdl,"Eqpdays",ylab="Log(OddsRatio of Churn)")
visreg(lrmdl,"Eqpdays",scale="response",ylab="Pr(Churn)")
# create a contour plot to visualize two effects at the same time (equipment days and months)
par(mfrow=c(1,1))
visreg2d(lrmdl,"Eqpdays","Months",plot.type="image",main="Log(OddsRatio of Adopt)")
visreg2d(lrmdl,"Eqpdays","Months",scale="response",plot.type="image",main="Pr(Adopt)")



#@clusterlr
###############################################################################
### cluster consumers based upon the data weighted by the logistic regression coefficients
### the purpose of this analysis is to identify a smaller set of 'prototypical' customers
### assumes that lrmdl has been estimated already (see Part1.R)
###############################################################################

# setup
if (!require(proxy)) {install.packages("proxy"); library(proxy)}


### cluster the contributions of the scores ###
# !! make sure lrmdl is set !!
# create predictions from the logistic regression model
userpred = predict(lrmdl,newdata=cell2cell,type='response')  # predict prob for all users

# create vector of variables used in model called mvarlist, and add other variables that we want to write out
# these lines require the lrmdl and ctree to be created appropriately above
varlist=names(coefficients(lrmdl))[-1]   # get the variables used in your logistic regression moodel, except the intercept which is in first position
varlist=unlist(unique(strsplit(varlist,":")))   # if you have interactions then you need to uncomment this line
print(varlist)  # vector of variables to save

# create new dataset with just selected variables for all users
userdata=model.matrix(lrmdl,data=cell2cell)  # construct the data used in the model

# "weight" the data using the logistic regression coefficients.  this allows the cluster to look at
# the variables based upon their contribution to their log-odds ratio
coefdata=summary(lrmdl)$coefficients  # extract coefficients estimates and std errors and z values
parm=coefdata[,1]  # just extract the parameter estimates in the 1st column
wuserdata=sweep(userdata,MARGIN=2,parm,"*")  # multiply each row in userdata by parm vector

# let's only consider those users that have a high probability of churn (say the top quartile)
# and place everyone else in a single cluster #0
p75=quantile(userpred,.75)
useridx=which(userpred>=p75)  # get the indices of those users that have high churn

# cluster the users into 20 groups
set.seed(612490)   # make sure we get the same solution
grpA=kmeans(wuserdata[useridx,],20)

# create a new cluster with everyone else (so we can work with the original dataset)
userdata = list()
userdata$cluster=rep(21,nrow(wuserdata))
userdata$cluster[useridx]=grpA$cluster

# describe the original data using the clusters
describeBy(cbind(userpred,cell2cell[,varlist]),group=userdata$cluster,fast=TRUE)

# create boxplot of userpred using the clusters
boxplot(userpred~userdata$cluster)
boxplot(cell2cell$Eqpdays~userdata$cluster)

### find prototypes of users in each cluster
# compute distances between data and centroids for grpA
# cdist is a matrix between original observation and each of the cluster centroids in the columns
# for example cdist[10,3] would be the distance between observation 10 in cell2cell and cluster #3
# need to use dist package from proxy since it can compute distances between variables
cdistA=proxy::dist(wuserdata,grpA$centers)
# find the closest observations to each centroid
cprototypeA=apply(cdistA,2,which.min)
# print the indices of the observations
print(cprototypeA)
# count the number of observations in each centroid
table(grpA$cluster)  # remember that we have taken just the top 25%, so the sum(table(grpA$cluster))=10000 not 40000
# print the prototypes for each centroid
print(cell2cell[cprototypeA,varlist])


### cluster the original cell2cell data ###

# alternative cluster the users into 20 groups just using raw data
set.seed(612390)   # make sure we get the same solution
grpB=kmeans(cell2cell[,varlist],20)

# describe the original data using the clusters
describeBy(cbind(userpred,cell2cell[,varlist]),group=grpB$cluster,fast=TRUE)

# create boxplot of userpred using the clusters
boxplot(userpred~grpB$cluster)
boxplot(cell2cell$Eqpdays~grpB$cluster)

# compare clusters
xtabs(~userdata$cluster+grpB$cluster)




#@whylrchurn
###############################################################################
### look for the three most likely reasons customer will churn
### ?future: subtract mean to give incremental contribution?
###############################################################################

# !! make sure lrmdl is set !!
# create predictions from the logistic regression model
userpred = predict(lrmdl,newdata=cell2cell,type='response')  # predict prob for all users

# create vector of variables used in model called mvarlist, and add other variables that we want to write out
# these lines require the lrmdl and ctree to be created appropriately above
varlist=names(coefficients(lrmdl))[-1]   # get the variables used in your logistic regression moodel, except the intercept which is in first position
varlist=unlist(unique(strsplit(varlist,":")))   # if you have interactions then you need to uncomment this line
print(varlist)  # vector of variables to save

# create new dataset with just selected variables for all users
userdata=model.matrix(lrmdl,data=cell2cell)  # construct the data used in the model

# "weight" the data using the logistic regression coefficients.  this allows the cluster to look at
# the variables based upon their contribution to their log-odds ratio
coefdata=summary(lrmdl)$coefficients  # extract coefficients estimates and std errors and z values
parm=coefdata[,1]  # just extract the parameter estimates in the 1st column
wuserdata=sweep(userdata,MARGIN=2,parm,"*")  # multiply each row in userdata by parm vector

# find the two most important reasons customer will churn
vnames=colnames(wuserdata[,-1])  # create vector of variables used in model (except for intercept)
result=t(apply(wuserdata[,-1],1,function(x) {vnames[order(x)[1:3]]}))  # sort effects by row and return top 3
result=as.data.frame(result)  # turn the output into a data frame
result$prob=userpred  # include the probability
head(result[result$prob>.75,])  # list out few users, the 0.75 is an arbitrary threshold -- change as desired -- perhaps to percentile
table(result$V1[result$prob>.75])  # reason for churn of most likely
table(result$V2[result$prob>.75])  # 2nd reason for churn
table(result$V3[result$prob>.75])  # 3rd reason for churn



#@prototypetree
###############################################################################
### find a prototypical user in each tree branch
###############################################################################

# setup
if (!require(proxy)) {install.packages("proxy"); library(proxy)}

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

# extract the notes associated with the trees
prnode=predict_nodes(ctree,cell2cell)
namenodes=unique(prnode)     # get the names of the nodes
numnodes=length(namenodes)   # count the number of terminal nodes (like the # of clusters)
table(prnode)
# print out the rules associated with the trees
result=path.rpart(ctree,nodes=prnode)  # output the rules for each node
print(result)   # print rules
#print(result$`99`)  # to look at a specific rule enter the rule number

# get list of variables
mvarlist.tree=names(ctree$variable.importance)   # if we want the list of variables in the tree then uncomment this line

# compute the centers for each cluster
centers=aggregate(cell2cell[,mvarlist.tree],list(prnode),mean)
# replace the mean with the mode for any factors
mvarlist.isfactor=sapply(cell2cell[1,mvarlist.tree],is.factor)
for (i in names(mvarlist.isfactor)[mvarlist.isfactor]) {
  tablemode=function(x) {names(table(x)[which.max(table(x))])}  # return a string with the modal value of an input factor
  centers[,i]=factor(aggregate(cell2cell[,i],list(prnode[]),tablemode)[,2],levels=levels(cell2cell[,i]))  # return factor using original levels
}

# for each leaf (final nodes in the tree) let's find the person that is closest to the center
prototypeTree=rep(NA,numnodes)  # create a place to store the nearest neighbor for each node
for (i in 1:numnodes) {
  # indexes of obserations in node
  nodeidx=which(prnode==namenodes[i])
  # compute distances to center of each cluster/node
  cdistnode=proxy::dist(cell2cell[nodeidx,mvarlist.tree],centers[i,mvarlist.tree])
  # find the closest observations to each centroid
  prototypeTree[i]=nodeidx[which.min(cdistnode)]
}
# print the indices of the observations
print(prototypeTree)
# count the number of observations in each centroid
table(prnode)  # remember that we have taken just the top 25%, so the sum(table(grpA$cluster))=10000 not 40000
# print the prototypes for each leaf/centroid
print(cell2cell[prototypeTree,mvarlist.tree])




#@randomforest
###############################################################################
### estimate a random forest
###
### uses randomForest package, but can also use cForest in party
### an advantage of cForest is that you can plot individual trees using prp
###############################################################################

# tools for random forest
if (!require(randomForest)) {install.packages("randomForest"); library(randomForest)} 

# estimate random forest
# !! change ntree=20 to larger values, but can take lots of time !!
# !! nodesize=30 limits branches to at least 30 observations !!
# caution: can take ~5 minutes to train using the following settings  (if time permits increase ntree)
rfmdl = randomForest(Churn~.,data=cell2cell[trainsample,],ntree=20,proximity=TRUE,importance=TRUE,nodesize=30)

# summary of forest
summary(rfmdl)
plot(rfmdl)        # plots the error rates per # of trees
varImpPlot(rfmdl)  # dot plot of the importance of the variables

# we can pull out specific trees from the forest
getTree(rfmdl,k=1,labelVar=TRUE)  # k can be any number between 1 and ntree

# compute predictions for the entire sample -- notice we only train on trainsample
pchurn.rf = predict(rfmdl,newdata=cell2cell,type='response')
cchurn.rf = (pchurn.rf>.5)+0   # make our predictions using a 50% cutoff, this can threshold can be changed
truechurn = cell2cell$Churn

# compute confusion matrix and some usual statistics (uncomment train and prediction samples if you want to see these statistics)
#confmatrix.summary(pchurn.rf[trainsample],cchurn.rf[trainsample],truechurn[trainsample])  # summary for training sample (look to see if train is substantially better than valid)
confmatrix.summary(pchurn.rf[validsample],cchurn.rf[validsample],truechurn[validsample])  # summary for validation sample
#confmatrix.summary(pchurn.rf[predsample],cchurn.rf[predsample],truechurn[predsample]) # summary for prediction sample (use this when you want to know how good your "final" model is)

# compute ROC and AUC
rocpred.rf = prediction(pchurn.rf[validsample],cell2cell$Churn[validsample])  # compute predictions using "prediction"
rocperf.rf = performance(rocpred.rf, measure = "tpr", x.measure = "fpr")
plot(rocperf.rf, col=rainbow(10)); abline(a=0, b= 1)
auc.tmp = performance(rocpred.rf,"auc")  # compute area under curve
(auc.rf = as.numeric(auc.tmp@y.values))

# plot response model
plotmo(rfmdl)             # evaluates selected input but holds other values at median
#plotmo(rfmdl,pmethod="apartdep")   # evaluates selected input and averages other values  !! pmethod="partdep" is better but slower !!

# understand relationship with inputs
( rfmdl.imp=importance(rfmdl) ) # larger values correspond with more important variables
impvar = rownames(rfmdl.imp)[order(rfmdl.imp[, 1], decreasing=TRUE)]
par(mfrow=c(2,3))  # put 2x3=6 plots in a panel
# !! change impvar[1:3] to impvar to see all variables
# !! change plotsample to trainsample to see all fitted data
for (i in seq_along(impvar[1:6])) {
  partialPlot(rfmdl, cell2cell[plotsample,], impvar[i], xlab=impvar[i],
              main=paste("Partial Dependence on", impvar[i]),ylim=c(0,1))
}

# we can compare the predicted responses vs actual response for each variable as well
# note: we can do this for any model
svarlist=rownames(importance(rfmdl)[1:6,])  # pick out the top six important variables
plotData <- lapply(names(cell2cell[,svarlist]), function(x){
  out <- data.frame(
    var = x,
    type = c(rep('Actual',nrow(cell2cell)),rep('Predicted',nrow(cell2cell))),
    value = c(cell2cell[,x],cell2cell[,x]),
    churn = c(cell2cell$Churn,pchurn.rf)
  )
  out$value <- out$value-min(out$value) #Normalize to [0,1]
  out$value <- out$value/max(out$value)
  out
})
plotData <- do.call(rbind,plotData)
qplot(value, churn, data=plotData, facets = type ~ var, geom='smooth', span = 0.01, ylim=c(0,1))



#@boostedtree
###############################################################################
### estimate a boosted tree with gbm
###
### some other alternatives are xgboost and LightGBM
### we use gbm since it has a nicer interface, but the others are newer and have
### computational advantages as well as other options
###############################################################################

# tools for gradient boosted tree
if (!require(gbm)) {install.packages("gbm"); library(gbm)} 

# estimate gradient boosted tree
# !! add cv.folds=5 for cross validation to find best # of trees, and n.cores=2 if you have multiple CPUs
# !! warning: may take >15 mins for n.trees=10000 (change value as time permits)
gbmdl.full=gbm(formula=Churn~.,data=cell2cell[trainsample,],
  distribution="bernoulli",n.trees=1000,shrinkage=.01,n.minobsinnode=20,cv.folds=5)

# if you have cv.folds>1 above then black line is training bernoulli deviance,
# green line is the testing bernoulli deviance, and best tree indicated by vertical blue line since
# it minimizes the testing error on the cross-validation folds
gbmdl.best=gbm.perf(gbmdl.full)
#gbmdl=gbmdl.best  # use this as your best gbm tree
gbmdl=gbmdl.full

# summary of forest
summary(gbmdl)

# visualize a tree (this is only one of many trees that make up gbmdl)
pretty.gbm.tree(gbmdl, i.tree=1)

# compute predictions for the entire sample -- notice we only train on trainsample
pchurn.gb = predict(gbmdl,newdata=cell2cell,type='response',n.trees=1000)
cchurn.gb = (pchurn.gb>.5)+0   # make our predictions using a 50% cutoff, this can threshold can be changed
truechurn = cell2cell$Churn

# compute confusion matrix and some usual statistics (uncomment train and prediction samples if you want to see these statistics)
#confmatrix.summary(pchurn.gb[trainsample],cchurn.gb[trainsample],truechurn[trainsample])  # summary for training sample (look to see if train is substantially better than valid)
confmatrix.summary(pchurn.gb[validsample],cchurn.gb[validsample],truechurn[validsample])  # summary for validation sample
#confmatrix.summary(pchurn.gb[predsample],cchurn.gb[predsample],truechurn[predsample]) # summary for prediction sample (use this when you want to know how good your "final" model is)

# compute ROC and AUC
rocpred.gb = prediction(pchurn.gb[validsample],cell2cell$Churn[validsample])  # compute predictions using "prediction"
rocperf.gb = performance(rocpred.gb, measure = "tpr", x.measure = "fpr")
plot(rocperf.gb, col=rainbow(10)); abline(a=0, b= 1)
auc.tmp = performance(rocpred.gb,"auc")  # compute area under curve
(auc.gb = as.numeric(auc.tmp@y.values))

# plot response model
plotmo(gbmdl)             # evaluates selected input but holds other values at median
#plotmo(gbmdl,pmethod="apartdep")   # evaluates selected input and averages other values  !! pmethod="partdep" is better but slower !!

# understand relationship with inputs
( gbmdl.imp=summary(gbmdl) ) # larger values correspond with more important variables
impvar=rownames(gbmdl.imp)[order(gbmdl.imp[, 2], decreasing=TRUE)]
impvar=impvar[1:6]  # just six most important

# we can compare the predicted responses vs actual response for each variable as well
# note: we can do this for any model
plotData <- lapply(names(cell2cell[,impvar]), function(x){
  out <- data.frame(
    var = x,
    type = c(rep('Actual',nrow(cell2cell)),rep('Predicted',nrow(cell2cell))),
    value = c(cell2cell[,x],cell2cell[,x]),
    churn = c(cell2cell$Churn,pchurn.gb)
  )
  out$value <- out$value-min(out$value) #Normalize to [0,1]
  out$value <- out$value/max(out$value)
  out
})
plotData <- do.call(rbind,plotData)
qplot(value, churn, data=plotData, facets = type ~ var, geom='smooth', span = 0.01, ylim=c(0,1))



#@compare
###############################################################################
### compare models using ROC plot
###############################################################################

# plot all ROC curves together
par(mfrow=c(1,1))
plot(rocperf.lr,col="red"); abline(a=0,b=1)
plot(rocperf.tree,add=TRUE,col="blue")
plot(rocperf.rf,add=TRUE,col="green")
plot(rocperf.gb,add=TRUE,col="violet")
legend("bottomright",c("LogRegr","Tree","Forest","GBM"),pch=15,col=c("red","blue","green","violet"),bty="n")



#@whatif
###############################################################################
### evaluate predictions for selected users for a variety of offers
### this section of code creates a number of alternative offers by changing
### the independent variables and then compute pr(churn) and LTV
###############################################################################

# setup information
irate=.05/12                  # annual discount rate
adja=.5/.02; adjb=.5/.98      # save the adjustments for a and b when computing adjusted churn
userlist=c(119,240,30,32)     # list of users to evaluate offers (!! change users if desired!!)
nuser=length(userlist)        # number of users
noffer=7                      # number of offers
if ("rpart" %in% class(mymodel)) {mtype='vector'} else {mtype='response'}  # set model type since predict needs different types for tree or logistic

# create data just for that user
cell2cellbase=cell2cell[userlist,]    # select users
cell2celluser=cell2cell[rep(userlist,each=noffer),]   # make a copy with all offers
cell2celluser$offer=rep(1:noffer,nuser)-1    # create another variable with the offers (start with 0)

# for each offer change the appropriate input values (!! change offers if desired !!)
offeridx=cell2celluser$offer                           # create index of offers
cell2celluser[offeridx==1,"Eqpdays"]=0      # new equipment
cell2celluser[offeridx==2,"Revenue"]=cell2celluser[cell2celluser$offer==2,"Revenue"]  # reduce price by $5
cell2celluser[offeridx==3,"Mou"]=0          # no usage
cell2celluser[offeridx==4,"Overage"]=0      # no overage
cell2celluser[offeridx==5,"Webcap"]=1       # web capable phone
cell2celluser[offeridx==6,"Actvsubs"]=2     # add another active subscriber

# evaluate the probability of churn for each offer
pchurn.unadj=predict(mymodel,newdata=cell2celluser,type=mtype)
# now adjust predictions to project probability of churn in the original data
a=pchurn.unadj/adja
b=(1-pchurn.unadj)/adjb
pchurn=a/(a+b)
# compute LTV
ltv.unadj=cell2celluser$Revenue*(1+irate)/(1+irate-(1-pchurn.unadj))
ltv=cell2celluser$Revenue*(1+irate)/(1+irate-(1-pchurn))   # still need to adjust for the cost of the offer

# bring all the columns together and combine with original cell2cell dataset
cell2cellbase$pchurn0=pchurn[offeridx==0]
cell2cellbase$pchurn1=pchurn[offeridx==1]
cell2cellbase$pchurn2=pchurn[offeridx==2]
cell2cellbase$pchurn3=pchurn[offeridx==3]
cell2cellbase$pchurn4=pchurn[offeridx==4]
cell2cellbase$pchurn5=pchurn[offeridx==5]
cell2cellbase$pchurn6=pchurn[offeridx==6]
cell2cellbase$ltv0=ltv[offeridx==0]
cell2cellbase$ltv1=ltv[offeridx==1]
cell2cellbase$ltv2=ltv[offeridx==2]
cell2cellbase$ltv3=ltv[offeridx==3]
cell2cellbase$ltv4=ltv[offeridx==4]
cell2cellbase$ltv5=ltv[offeridx==5]
cell2cellbase$ltv6=ltv[offeridx==6]

# look at the results
head(cell2cellbase)
describe(cell2cellbase)
