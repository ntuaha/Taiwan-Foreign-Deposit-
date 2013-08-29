rm(list=ls(all=TRUE))
library(ggplot2)

d = read.table('/Users/aha/Dropbox/Project/Financial/Codes/csv/Total.csv',header=T,sep=",",encoding="UTF-8")
names(d) = c("date","ALL_MY","ALL_FY","ALL_Y","DB_MY","DB_FY","DB_Y","OS_MY","OS_FY","OS_Y")






p = ggplot(data=d,aes(x=1:length(ALL_Y),y=ALL_Y))
p+geom_line()+stat_smooth()