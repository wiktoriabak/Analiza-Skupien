library(dplyr)
library(readxl)
library(tibble)
library(outliers)  
library(ggplot2)
library(cluster)  
library(reshape2)
library(fmsb)  
library(e1071)
library(psych)
library(dplyr)
library(corrplot)

pierw_dane <- read_excel("C:/Users/wiki2/OneDrive/Pulpit/Dane.xlsx", sheet = 1)
pierw_dane <- column_to_rownames(pierw_dane, var = names(pierw_dane)[1])
dane <- read_excel("C:/Users/wiki2/OneDrive/Pulpit/data.xlsx", sheet = 1)
dane <- column_to_rownames(dane, var = names(dane)[1])

dane_stand <- scale(dane)


cv <- apply(pierw_dane, 2, function(x) (sd(x, na.rm = TRUE) / mean(x, na.rm = TRUE)))

tabela_statystyk <- data.frame(
  Średnia = round(apply(pierw_dane, 2, function(x) mean(x, na.rm = TRUE)), 3),
  Odchylenie_standardowe = round(apply(dane, 2, function(x) sd(x, na.rm = TRUE)), 3),
  Minimum = round(apply(pierw_dane, 2, function(x) min(x, na.rm = TRUE)), 3),
  Maksimum = round(apply(pierw_dane, 2, function(x) max(x, na.rm = TRUE)), 3),
  Skośność = round(apply(pierw_dane, 2, function(x) skew(x, na.rm = TRUE)), 3),
  Kurtoza = round(apply(pierw_dane, 2, function(x) kurtosi(x, na.rm = TRUE)), 3),
  Wsp_zmiennosci = round(cv, 3)
)

write.csv(tabela_statystyk, "C:/Users/wiki2/OneDrive/Pulpit/eda_results.csv", row.names = FALSE)


correlation_matrix <- cor(pierw_dane, use = "complete.obs")


corrplot(correlation_matrix, method = "color", tl.cex = 0.9,
         col = colorRampPalette(c("#3182bd", "#f7f7f7", "#fc4e2a"))(200),
         tl.col = "black",
         addgrid.col = "black",
         addCoef.col = "black") 


pierw_dane_long <- melt(pierw_dane)

zmienne <- unique(pierw_dane_long$variable)

zmienne_1 <- zmienne[1:6]
zmienne_2 <- zmienne[7:length(zmienne)]

plot1 <- ggplot(pierw_dane_long[pierw_dane_long$variable %in% zmienne_1, ], aes(x = variable, y = value)) +
  geom_boxplot() +
  theme_minimal() +
  labs(x = "Zmienna", y = "Wartość") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1),
        strip.text = element_text(face = "bold")) +  
  facet_wrap(~variable, scales = "free", ncol = 3) +
  scale_x_discrete(labels = NULL)  

plot2 <- ggplot(pierw_dane_long[pierw_dane_long$variable %in% zmienne_2, ], aes(x = variable, y = value)) +
  geom_boxplot() +
  theme_minimal() +
  labs(x = "Zmienna", y = "Wartość") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1),
        strip.text = element_text(face = "bold")) +  
  facet_wrap(~variable, scales = "free", ncol = 3) +
  scale_x_discrete(labels = NULL) 

print(plot1)
print(plot2)


prepare_radar_data <- function(data, selected_index) {
  
  min_vals <- apply(data, 2, min)
  max_vals <- apply(data, 2, max)
  
  radar_data <- rbind(max_vals, min_vals, data[selected_index, ])
  
  radar_data <- as.data.frame(radar_data)
  
  return(radar_data)
}


plot_radar_chart <- function(data, title) {
  
  radarchart(data, 
             axistype = 0,       
             pcol = "black",                
             pfcol = rgb(0, 0, 0, 0.7),       
             plwd = 2,        
             cglcol = "grey",     
             cglty = 1,        
             axislabcol = "black",    
             cglwd = 0.8,
             vlcex = 0.8            
  )
  
  title(main = title, cex.main = 1.5)
}


radar_data <- prepare_radar_data(dane_stand, 34)

plot_radar_chart(radar_data, rownames(dane)[34])


for (i in 1:ncol(dane_stand)) {
  Q1 <- quantile(dane_stand[, i], 0.25)
  Q3 <- quantile(dane_stand[, i], 0.75)
  IQR <- Q3 - Q1
  
  dolny_was <- Q1 - 1.5 * IQR
  gorny_was <- Q3 + 1.5 * IQR
  
  dane_stand[dane_stand[, i] < dolny_was, i] <- dolny_was
  dane_stand[dane_stand[, i] > gorny_was, i] <- gorny_was
}

przekraczajace_3 <- which(abs(dane_stand) > 3, arr.ind = TRUE)

if (nrow(przekraczajace_3) > 0) {
  cat("Wartości przekraczające 3 po korekcie:\n")
  print(przekraczajace_3)
} else {
  cat("Brak wartości przekraczających 3 po korekcie.\n")
}


suma_standaryzowana <- rowSums(dane_stand)

wzorzec <- apply(dane_stand, 2, max)
odleglosc_hellwiga <- apply(dane_stand, 1, function(x) sqrt(sum((x - wzorzec)^2)))
hellwig <- 1 - (odleglosc_hellwiga / max(odleglosc_hellwiga))

rangi <- apply(dane_stand, 2, rank)
suma_rang <- rowSums(rangi)


wyniki <- data.frame(
  Powiat = rownames(dane),
  Suma_Standaryzowana = suma_standaryzowana,
  Hellwig = hellwig,
  Suma_Rang = suma_rang
)

wyniki <- wyniki %>%
  mutate(
    Miejsce_Suma_Standaryzowana = rank(-Suma_Standaryzowana),  
    Miejsce_Hellwig = rank(-Hellwig),                         
    Miejsce_Suma_Rang = rank(-Suma_Rang)                       
  )

wyniki <- wyniki %>%
  mutate(Zbiorcze_Miejsce = Miejsce_Suma_Standaryzowana + Miejsce_Hellwig + Miejsce_Suma_Rang) %>%
  arrange(Zbiorcze_Miejsce) %>%
  mutate(Zbiorczy_Ranking = rank(Zbiorcze_Miejsce))

print(wyniki)



suma_standaryzowana_standardized <- (suma_standaryzowana - min(suma_standaryzowana)) / (max(suma_standaryzowana) - min(suma_standaryzowana))

standardized_results <- data.frame(
  Powiat = rownames(dane),
  Suma_Standaryzowana_Standardized = suma_standaryzowana_standardized
)

print(standardized_results)

write.csv(standardized_results, "C:/Users/wiki2/OneDrive/Pulpit/standardized_sums.csv", row.names = FALSE)


classify_group <- function(value, mean_value, sd_value) {
  if (value >= mean_value + sd_value) {
    return("I")
  } else if (value >= mean_value && value < mean_value + sd_value) {
    return("II")
  } else if (value >= mean_value - sd_value && value < mean_value) {
    return("III")
  } else {
    return("IV")
  }
}

wyniki <- wyniki %>%
  mutate(
    Grupa_Suma_Standaryzowana = sapply(Suma_Standaryzowana, classify_group, mean(Suma_Standaryzowana), sd(Suma_Standaryzowana)),
    Grupa_Hellwig = sapply(Hellwig, classify_group, mean(Hellwig), sd(Hellwig)),
    Grupa_Suma_Rang = sapply(Suma_Rang, classify_group, mean(Suma_Rang), sd(Suma_Rang))
  )


print(wyniki)

write.csv(wyniki, "C:/Users/wiki2/OneDrive/Pulpit/porzadkowanie_liniowe.csv", row.names = FALSE)

korelacja_kendall <- cor(wyniki[, c("Suma_Standaryzowana", "Hellwig", "Suma_Rang")], method = "kendall")
print(korelacja_kendall)


results <- list()

for (k in 2:10) {
  kmeans_result <- kmeans(dane_stand, centers = k)
  results[[as.character(k)]] <- kmeans_result
}

k_opt <- 3
kmeans_result <- results[[as.character(k_opt)]]

pca_result <- prcomp(dane_stand, center = TRUE, scale. = TRUE)
pca_data <- as.data.frame(pca_result$x)

pca_data$Klaster <- as.factor(kmeans_result$cluster)


ggplot(pca_data, aes(x = PC1, y = PC2, color = Klaster, shape = Klaster)) +
  geom_point(size = 4, alpha = 0.8, stroke = 1.2) +  # Zwiększenie rozmiaru punktów i dodanie konturów
  labs(title = paste("Wizualizacja klastrów k-średnich dla k =", k_opt),
       x = "Główna składowa 1 (PC1)",
       y = "Główna składowa 2 (PC2)") +
  theme_minimal() +
  scale_color_brewer(palette = "Set1") +
  theme(legend.position = "bottom")


distan <- dist(dane_stand, method = "euclidean")

hierarchical_result <- hclust(distan, method = "ward.D2")

plot(hierarchical_result, main = "Dendrogram dla grupowania hierarchicznego (metoda Warda)",
     xlab = "Obiekty", ylab = "Odległość", sub = "", cex = 0.5) 

rect.hclust(hierarchical_result, k = 5, border = "red")


k_opt <- 5
clusters <- cutree(hierarchical_result, k = k_opt)


pca_data$Klaster_hierarchical <- as.factor(clusters)


clusters_kmeans_assignment <- data.frame(
  Powiat = rownames(dane),
  Klaster_KMeans = kmeans_result$cluster
)

clusters_kmeans_assignment_sorted <- clusters_kmeans_assignment %>%
  arrange(Klaster_KMeans)

print(clusters_kmeans_assignment_sorted)



clusters_assignment <- data.frame(
  Powiat = rownames(dane),
  Klaster_Hierarchical = clusters
)

clusters_assignment_sorted <- clusters_assignment %>%
  arrange(Klaster_Hierarchical)

print(clusters_assignment_sorted)



dane_kmeans <- as.data.frame(pierw_dane)
dane_kmeans$Cluster_KMeans <- kmeans_result$cluster

kmeans_stats <- dane_kmeans %>%
  group_by(Cluster_KMeans) %>%
  summarise_all(list(mean = ~mean(.), sd = ~sd(.)))

print(kmeans_stats)

write.csv(kmeans_stats, "C:/Users/wiki2/OneDrive/Pulpit/kmeans_stats.csv", row.names = FALSE)



dane_hierarchical <- as.data.frame(pierw_dane)
dane_hierarchical$Cluster_Hierarchical <- clusters

hierarchical_stats <- dane_hierarchical %>%
  group_by(Cluster_Hierarchical) %>%
  summarise_all(list(mean = ~mean(.), sd = ~sd(.)))

print(hierarchical_stats)

write.csv(hierarchical_stats, "C:/Users/wiki2/OneDrive/Pulpit/hierarchical_stats.csv", row.names = FALSE)
