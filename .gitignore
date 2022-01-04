#Load Package
library(tm)
library(stringr)
library(dplyr)
library(katadasaR)
library(openxlsx)
library(tokenizers)

#Memasukkan Data 
setwd("D:/RUNNING")
data_laporan<-read.xlsx("data.xlsx")

#Membentuk Corpus 
corpusdoc=VCorpus(VectorSource(data_laporan$LAPORAN))

#Case Folding
doc_casefolding<-tm_map(corpusdoc, content_transformer(tolower))
prep_tolower<-data.frame(text=sapply(doc_casefolding,as.character,stringAsFactors=FALSE))
write.csv(prep_tolower,file="prep_tolower.csv")

#Remove URL
removeURL<-function(x)gsub("http[^[:space:]]*","",x)
doc_URL<-tm_map(doc_casefolding, content_transformer(removeURL))
prep_removeurl<-data.frame(text=sapply(doc_URL,as.character,stringAsFactors=FALSE))
write.csv(prep_removeurl,file="prep_removeurl.csv")

#Remove Mention
removeMention <- function(x) gsub("@\\w+", "", x)
doc_mention<- tm_map(doc_URL,content_transformer(removeMention))
prep_removemention <- data.frame(text=sapply(doc_mention,as.character), stringsAsFactors=FALSE)
write.csv(prep_removemention,file="prep_removemention.csv")



#Remove Hastag
remove.hastag<-function(x)gsub("#\\S+","",x)
doc_hastag<-tm_map(doc_mention,content_transformer(remove.hastag))
prep_removehastag <- data.frame(text=sapply(doc_hastag,as.character), stringsAsFactors=FALSE)
write.csv(prep_removehastag,file="prep_removehastag.csv")

#	Remove Punctuation
removepunct <- function(x) gsub("[[:punct:]]"," ",x)
doc_punctuation<- tm_map(doc_hastag,content_transformer(removepunct))
prep_removepunctuation<-data.frame(text=sapply(doc_punctuation,as.character,stringAsFactors=FALSE))
write.csv(prep_removepunctuation,file="prep_removepunct.csv")

#Remove Number
doc_nonnumber<-tm_map(doc_punctuation, content_transformer(removeNumbers))
prep_removenumber<-data.frame(text=sapply(doc_nonnumber,as.character,stringAsFactors=FALSE))
write.csv(prep_removenumber,file="prep_removenumber.csv")

#Pembakuan Kata
spell.lex <- read.xlsx("convertword.xlsx")
spell.correction <- content_transformer(function(x, dict){
  words <- sapply(unlist(strsplit(x, "\\s+")),function(x){
    if(is.na(spell.lex[match(x, dict$spell),"word"])){
      x <- x
    } else{
      x <- spell.lex[match(x, dict$spell),"word"]
    }
  })
  x <- paste(words, collapse = " ")
})
doc_normalisasi <- tm_map(doc_nonnumber, spell.correction, spell.lex)
prep_normalisasi <- data.frame(text=sapply(doc_normalisasi, as.character),stringsAsFactors=FALSE) 
write.csv(prep_normalisasi,file="prep_normalisasi.csv")

#Stopword Removal
cStopwordID<-readLines("Stopword.csv")
doc_stopword<-tm_map(doc_normalisasi, removeWords, cStopwordID)
prep_stopword <- data.frame(text=sapply(doc_stopword, as.character),stringsAsFactors=FALSE) 
write.csv(prep_stopword,file="prep_stopword.csv")

#Strip Whitespace
doc_whitespace <- tm_map(doc_stopword,stripWhitespace)
prep_whitespace<-data.frame(text=sapply(doc_whitespace, as.character), stringsAsFactors = FALSE)
write.csv(prep_whitespace, file = "prep_whitespacecoba.csv")

#	Stemming
stemming=function(x){
  paste(sapply(unlist(str_split(x,'\\s+')),katadasaR), collapse=" ")
}
doc_stemming<-tm_map(doc_whitespace,content_transformer(stemming))
prep_stemming<-data.frame(text=sapply(doc_stemming, as.character), stringsAsFactors = FALSE)
write.csv(prep_stemming, file = "prep_stemming.csv")
 
#Pembobotan TF-IDF

#Pembobotan TF
dtm.tf <- DocumentTermMatrix(doc_stemming)
dtm.tf.matriks = as.matrix(dtm.tf)
write.csv(dtm.tf.matriks, file = "matriksTF.csv")

#Pembobotan TF-IDF
dtm.tf.idf <- weightTfIdf(m = dtm.tf, normalize = TRUE)
dtm.tf.idf.matriks=as.matrix(dtm.tf.idf)
write.csv(dtm.tf.idf.matriks, file = "matriksTF-IDF.csv")
 
#Fuzzy C-Means Clustering
#Load Package
library(ppclust)
library(cluster)

#Proses Fuzzy C-Means Clustering
#Cluster 3
u3<- inaparc::imembrand(nrow(dtm.tf.idf.matriks), k=3)$u
member.inisial3<-as.data.frame(u3)
write.csv(member.inisial3,file="member.inisial3.csv")
v3<- inaparc::kmpp(dtm.tf.idf.matriks, k=3)$v
center.inisial3<-as.data.frame(v3)
write.csv(center.inisial3,file="center.inisial3.csv")
res.fcm3<-fcm(dtm.tf.idf.matriks, centers=v3,memberships = u3, m=2)
output<- data.frame(data_laporan$'NO', res.fcm3$cluster, res.fcm3$u)
write.csv(output,file='output.csv')
#Cluster 4
u4<- inaparc::imembrand(nrow(dtm.tf.idf.matriks), k=4)$u
member.inisial4<-as.data.frame(u4)
write.csv(member.inisial4,file="member.inisial4.csv")
v3<- inaparc::kmpp(dtm.tf.idf.matriks, k=4)$v
center.inisial4<-as.data.frame(v4)
write.csv(center.inisial4,file="center.inisial4.csv")
res.fcm4<-fcm(dtm.tf.idf.matriks, centers=v4,memberships = u4, m=2)
#Cluster 5
u5<- inaparc::imembrand(nrow(dtm.tf.idf.matriks), k=5)$u
member.inisial5<-as.data.frame(u5)
write.csv(member.inisial5,file="member.inisial5.csv")
v5<- inaparc::kmpp(dtm.tf.idf.matriks, k=5)$v
center.inisial5<-as.data.frame(v5)
write.csv(center.inisial5,file="center.inisial5.csv")
res.fcm5<-fcm(dtm.tf.idf.matriks, centers=v5,memberships = u5, m=2)
#Cluster 6
u6<- inaparc::imembrand(nrow(dtm.tf.idf.matriks), k=6)$u
member.inisial6<-as.data.frame(u6)
write.csv(member.inisial6,file="member.inisial6.csv")
v6<- inaparc::kmpp(dtm.tf.idf.matriks, k=3)$v
center.inisial6<-as.data.frame(v6)
write.csv(center.inisial6,file="center.inisial6.csv")
res.fcm6<-fcm(dtm.tf.idf.matriks, centers=v6,memberships = u6, m=2)
#Cluster 7
u7<- inaparc::imembrand(nrow(dtm.tf.idf.matriks), k=7)$u
member.inisial7<-as.data.frame(u7)
write.csv(member.inisial7,file="member.inisial7.csv")
v7<- inaparc::kmpp(dtm.tf.idf.matriks, k=7)$v
center.inisial7<-as.data.frame(v7)
write.csv(center.inisial7,file="center.inisial7.csv")
res.fcm7<-fcm(dtm.tf.idf.matriks, centers=v7,memberships = u7, m=2)

 
# Silhouette Coefficient
#Load Package
library(fclust)
#Proses silhouette Coefficient
#Cluster 3
res.fcm33<-ppclust3(res.fcm3,"fclust")
idxs3<-SIL(res.fcm33$Xca,res.fcm33$U)
idxs3$sil
#Cluster 4
res.fcm44<-ppclust2(res.fcm4,"fclust")
idxs4<-SIL(res.fcm44$Xca,res.fcm44$U)
idxs4$sil
#Cluster 5
res.fcm55<-ppclust2(res.fcm5,"fclust")
idxs5<-SIL(res.fcm22$Xca,res.fcm55$U)
idxs5$sil
#Cluster 6
res.fcm66<-ppclust2(res.fcm6,"fclust")
idxs6<-SIL(res.fcm66$Xca,res.fcm66$U)
idxs6$sil
#Cluster 7
res.fcm77<-ppclust2(res.fcm7,"fclust")
idxs7<-SIL(res.fcm77$Xca,res.fcm77$U)
idxs7$sil

 
#Pembentukan Wordcloud
#Load Package
library(tokenizers)
library(quanteda)
library(wordcloud2)
# install webshot
webshot::install_phantomjs()
library("htmlwidgets")

#Pembentukan Wordcloud
cluster.list3 <- as.data.frame(res.fcm3$cluster)
write.csv(cluster.list3, file = "cluster_list3.csv")
doc<-data.frame(text=sapply(doc_stemming, as.character), stringsAsFactors = FALSE)
doc.cluster <- cbind(doc, cluster.list3)
names(doc.cluster) <- c("text","cluster")
str(doc.cluster)
#wordcloud of cluster pertama
doc.cluster1<-doc.cluster[which(doc.cluster$cluster==1),]
tokenize1<-tokenize_words(doc.cluster1$text)
tokens.tfidf1 <- tokens(tokenize1, what = "word", 
                        remove_numbers = TRUE, remove_punct = TRUE,
                        remove_symbols = TRUE,  remove_url = TRUE,
                        remove_separators = TRUE)
doc.c1 <- as.character(tokens.tfidf1)
doc1 <- table(doc.c1) %>% 
  as.data.frame() %>% 
  arrange(desc(Freq))
names(doc1) <- c("word","Freq")
head(doc1)
graph1 <- wordcloud2(doc1, size = 1, minSize = 0, gridSize = 0,
                     fontFamily = 'Segoe UI', fontWeight = 'bold',
                     color = 'random-dark', backgroundColor = "white",
                     minRotation = -pi/4, maxRotation = pi/4, shuffle = TRUE,
                     rotateRatio = 0.4, shape = 'circle', ellipticity = 0.65,
                     widgetsize = NULL, figPath = NULL, hoverFunction = NULL)
saveWidget(graph1,"tmp1.html",selfcontained = F)
webshot("tmp1.html","fig_1.pdf", delay =5, vwidth = 650, vheight=650)
#wordcloud of cluster kedua
doc.cluster2<-doc.cluster[which(doc.cluster$cluster==2),]
tokenize2<-tokenize_words(doc.cluster2$text)
tokens.tfidf2 <- tokens(tokenize2, what = "word", 
                        remove_numbers = TRUE, remove_punct = TRUE,
                        remove_symbols = TRUE,  remove_url = TRUE,
                        remove_separators = TRUE)
doc.c2 <- as.character(tokens.tfidf2)
doc2 <- table(doc.c2) %>% 
  as.data.frame() %>% 
  arrange(desc(Freq))
names(doc2) <- c("word","Freq")
graph2 <- wordcloud2(doc2, size = 1, minSize = 0, gridSize = 0,
                     fontFamily = 'Segoe UI', fontWeight = 'bold',
                     color = 'random-dark', backgroundColor = "white",
                     minRotation = -pi/4, maxRotation = pi/4, shuffle = TRUE,
                     rotateRatio = 0.4, shape = 'circle', ellipticity = 0.65,
                     widgetsize = NULL, figPath = NULL, hoverFunction = NULL)
saveWidget(graph2,"tmp2.html",selfcontained = F)
webshot("tmp2.html","fig_2.pdf", delay =5, vwidth = 650, vheight=650)
#wordcloud of cluster ketiga
doc.cluster3<-doc.cluster[which(doc.cluster$cluster==3),]
tokenize3<-tokenize_words(doc.cluster3$text)
tokens.tfidf3 <- tokens(tokenize3, what = "word", 
                        remove_numbers = TRUE, remove_punct = TRUE,
                        remove_symbols = TRUE,  remove_url = TRUE,
                        remove_separators = TRUE)
doc.c3 <- as.character(tokens.tfidf3)
doc3 <- table(doc.c3) %>% 
  as.data.frame() %>% 
  arrange(desc(Freq))
names(doc3) <- c("word","Freq")
head(doc3)
graph3 <- wordcloud2(doc3, size = 1, minSize = 0, gridSize = 0,
                     fontFamily = 'Segoe UI', fontWeight = 'bold',
                     color = 'random-dark', backgroundColor = "white",
                     minRotation = -pi/4, maxRotation = pi/4, shuffle = TRUE,
                     rotateRatio = 0.4, shape = 'circle', ellipticity = 0.65,
                     widgetsize = NULL, figPath = NULL, hoverFunction = NULL)
saveWidget(graph3,"tmp3.html",selfcontained = F)
webshot("tmp3.html","fig_3.pdf", delay =5, vwidth = 650, vheight=650)
