library(readxl)
library(dplyr)
library(writexl)
library(openxlsx)

#=============================================
#UNIT 2 & 3 (Integrasi dan Mengumpulkan Data)
#=============================================
#Integrasi data
datapersonal = read_excel("Dataset_Portofolio1.xlsx", sheet = 'Data_Personal')
datagaji = read_excel('Dataset_Portofolio1.xlsx', sheet = 'Data_Gaji')
databenefit = read_excel('Dataset_Portofolio1.xlsx', sheet = 'Data_Benefit')

colnames(datapersonal)
colnames(datagaji)
colnames(databenefit)

View(datapersonal)
View(datagaji)
View(databenefit)

#Integrasi dan Pengumpulan Data Data Personal dan Data Gaji berdasarkan Employee Id & Year
di1 = datapersonal %>% left_join (datagaji, by = c('Employee Identifier','Year'))
View(di1)
#Integrasi dan pengumpulan Data di1 dengan Data Benefit berdasarkan Employee Id & Year
di2 = di1 %>% left_join (databenefit, by = c('Employee Identifier','Year'))
View(di2)

file_datamentah = "Dataset_Portofolio1.xlsx"
wbgaji = loadWorkbook(file_datamentah)
addWorksheet(wbgaji, sheetName = 'Data Integrasi')
writeData(wbgaji, sheet = "Data Integrasi", x=di2)
saveWorkbook(wbgaji, file = "Dataset_Portofolio1.xlsx", overwrite = TRUE)

#============================================
#UNIT 4 & 5 (Menelaah dan Memvalidasi Data)
#============================================
# === Menelaah Data === #
str(di2)
summary(di2) #Masih terdapat data negatif, jadikan absolut

# === Validasi Data === #
#1.Tidak ada data negatif (data absolut)
#2. Tidak ada data kosong (N.A)
#3.Job code dan Nama Job berjumlah sama
di2abs = di2 %>%
  mutate(across(where(is.numeric),abs))
summary(di2abs)
print(di2abs, n=23)

colSums(is.na(di2))

#Jumlah data duplikat
sum(duplicated(di2[c('Employee Identifier', 'Year')]))
# Memastikan Jumlah Code dan Nama Job harus sama
total_job_code <- n_distinct(di2$`Job Code`)
cat("Total Job Code:", total_job_code, "\n")

total_job <- n_distinct(di2$`Job`)
cat("Total Job:", total_job, "\n")

# Cari job code yang memiliki lebih dari 1 nama jabatan
job_code_duplikat = di2 %>%
  group_by(`Job Code`) %>%
  summarise(
    Jumlah_Nama = n_distinct(`Job`),
    Daftar_Nama = paste(unique(`Job`), collapse = " | ")
  ) %>%
  filter(Jumlah_Nama > 1) %>%
  arrange(desc(Jumlah_Nama))
print(job_code_duplikat) #Ubah melalui Excel

`# Jumlah kode jabatan dan nama job
total_jcduplikat= n_distinct(di2$"Job Code")
cat("Total Job Code:", total_jcduplikat, "\n")

total_job <- n_distinct(di2$"Job")
cat("Total Job:", total_job, "\n")




# ====================================================
# UNIT 6 & 7 (Menentukan Objek dan Membersihkan Data)
# ====================================================
# === Menentukan Objek === #
View(di2abs)
MainData = di2abs [,-c(1,3,4,5,6,7,8,9,10)]
colnames(MainData)
str(MainData)
summary(MainData)
View(MainData)


# === Pembersihan Data === #
# Hapus duplikat berdasarkan (Employee Identifier + Year)
MD_clean <- MainData %>% 
  distinct(`Employee Identifier`, Year, .keep_all = TRUE)
# Cek hasil
cat("Sebelum:", nrow(MainData), "\n")
cat("Sesudah:", nrow(MD_clean), "\n")
cat("Duplikat terhapus:", nrow(MainData) - nrow(MD_clean))
View(MD_clean)


#==============================
# UNIT 8 Business Intelligence
#==============================
# === Analis Data Project === #
MD_Proj <- MD_clean %>%
  group_by(`Job Code`, `Job`, Year) %>%
  summarise(
    #Komponen Salary
    Totals = sum(`Total Salary`, na.rm = TRUE),
    Salaries = sum(Salaries, na.rm = TRUE),
    Overtime = sum(Overtime, na.rm = TRUE),
    Others = sum(`Other Salaries`, na.rm = TRUE),
    #Komponen Benefit
    Totalb = sum(`Total Benefits`, na.rm = TRUE),
    Retirement = sum(Retirement, na.rm = TRUE),
    Health = sum(`Health/Dental`, na.rm = TRUE),
    Otherb = sum(`Other Benefits`, na.rm = TRUE),
    #Keseluruhan Belanja Pegawai
    Totalc = sum(`Total Compensation`, na.rm = TRUE),
    #Tambahan
    Jumlah_Pegawai = n_distinct(`Employee Identifier`),
    Rata2_Gaji = mean(`Total Salary`, na.rm = TRUE),
    Rata2_Benefit = mean(`Total Benefits`, na.rm = TRUE),
    Rata2_Kompensasi = mean(`Total Compensation`, na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  mutate(
    Proporsi_Gaji = Totals/ Totalc * 100,
    Proporsi_Benefit = Totalb / Totalc * 100
  ) %>%
  arrange(Year) #Ascending berrdasarkan tahun dari 2013
View(MD_Proj)


# +++++++++++++++++++++++++++++++++++++++++++++ #
#     Tren Belanja Pegawai per Job (Case 1)
# +++++++++++++++++++++++++++++++++++++++++++++ #
tren_job = MD_clean %>%
  group_by(`Job Code`, `Job`, Year) %>%
  summarise(
    Total_Kompensasi = sum(`Total Compensation`, na.rm = TRUE),
    Jumlah_Pegawai = n_distinct(`Employee Identifier`),
    Rata2_Kompensasi = mean(`Total Compensation`, na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  arrange(`Job Code` , Year)
View(tren_job)

library(tidyr)
trenjob_tabel = tren_job %>%
  pivot_wider(
    id_cols = c(`Job Code`, `Job`),
    names_from = Year,
    values_from = c(Total_Kompensasi, Jumlah_Pegawai, Rata2_Kompensasi),
    names_glue = "{Year}_{.value}")

tahun_list = sort(unique(tren_job$Year))

trenjob_tabel = trenjob_tabel %>%
  mutate(
    Total_3Tahun = rowSums(select(., starts_with("201") & ends_with("_Total_Kompensasi")), na.rm = TRUE),
    Rata2_3Tahun = Total_3Tahun / length(tahun_list)
  ) %>%
  relocate(Total_3Tahun, .after = `Job`) %>%
  relocate(Rata2_3Tahun, .after = Total_3Tahun)

trenjob_tabel = trenjob_tabel %>%
  arrange(desc(Total_3Tahun))

kolom_urut <- c("Job Code", "Job")
for(tahun in tahun_list) {
  kolom_urut = c(kolom_urut, 
                  paste0(tahun, "_Total_Kompensasi"),
                  paste0(tahun, "_Jumlah_Pegawai"),
                  paste0(tahun, "_Rata2_Kompensasi"))
}

trenjob_tabel = trenjob_tabel[, kolom_urut]
View(trenjob_tabel)


# === Visualisasi Case 1 === #
library(ggplot2)
# Ambil top 5 job dengan total kompensasi tertinggi
top5_job = tren_job %>%
  group_by(`Job Code`) %>%
  summarise(Total = sum(Total_Kompensasi)) %>%
  arrange(desc(Total)) %>%
  head(5) %>%
  pull(`Job Code`)

# Filter tren untuk top 5
tren_top5 = tren_job %>%
  filter(`Job Code` %in% top5_job)
View(tren_top5)
# Visualisasi tren
plot_tren = ggplot(tren_top5, aes(x = factor(Year), 
                                   y = Total_Kompensasi/1000, 
                                   color = `Job`, 
                                   group = `Job`)) +
  geom_line(linewidth = 1.5) +
  geom_point(size = 4) +
  geom_text(aes(label = round(Total_Kompensasi/1000, 2)), 
            vjust = -1, size = 3) +
  labs(
    title = "Tren Belanja Pegawai Top 5 Job (3 Tahun)",
    subtitle = "Total Kompensasi per Tahun (USD)",
    x = "Tahun",
    y = "Total Kompensasi (USD)",
    color = "Jabatan"
  ) +
  theme_minimal() +
  theme(
    legend.position = "bottom",
    plot.title = element_text(size = 14, face = "bold"),
    axis.text.x = element_text(size = 12)
  )
print(plot_tren)



# +++++++++++++++++++++++++++++++++++++++++++++++++++++ #
#     Pegawai dengan Multiple Job Tahunan (Case 2)
# +++++++++++++++++++++++++++++++++++++++++++++++++++++ #

# Cek employee yang memiliki lebih dari 1 job (per tahun)
employee_multiple_job = MainData %>%
  group_by(`Employee Identifier`, Year) %>%
  summarise(
    Jumlah_Job = n_distinct(`Job Code`),
    Daftar_Job = paste(unique(`Job`), collapse = " | "),
    Daftar_Job_Code = paste(unique(`Job Code`), collapse = " | "),
    .groups = 'drop'
  ) %>%
  filter(Jumlah_Job > 1) %>%
  arrange(desc(Jumlah_Job), `Employee Identifier`, Year)

print("===== EMPLOYEE DENGAN MULTIPLE JOB =====")
print(employee_multiple_job)
View(employee_multiple_job)


# Statistik multiple job
stat_multiple_job = data.frame(
  Metrik = c(
    "Total Employee Unik",
    "Employee dengan Multiple Job",
    "Persentase Multiple Job"
  ),
  Nilai = c(
    n_distinct(MD_clean$`Employee Identifier`),
    n_distinct(employee_multiple_job$`Employee Identifier`),
    paste0(round(n_distinct(employee_multiple_job$`Employee Identifier`) / 
                   n_distinct(MD_clean$`Employee Identifier`) * 100, 2), "%")
  )
)

print("===== STATISTIK MULTIPLE JOB =====")
print(stat_multiple_job)

# === Visualisasi Case 3 === #
proporsi_multiple = data.frame(
  Kategori = c("Multiple Job", "Single Job"),
  Jumlah = c(
    n_distinct(employee_multiple_job$`Employee Identifier`),
    n_distinct(MD_clean$`Employee Identifier`) - 
      n_distinct(employee_multiple_job$`Employee Identifier`)
  )
)

plot_pie = ggplot(proporsi_multiple, aes(x = "", y = Jumlah, fill = Kategori)) +
  geom_bar(stat = "identity", width = 1) +
  coord_polar("y", start = 0) +
  geom_text(aes(label = paste0(Jumlah, "\n(", round(Jumlah/sum(Jumlah)*100, 1), "%)")),
            position = position_stack(vjust = 0.5), size = 5) +
  labs(
    title = "Proporsi Employee dengan Multiple Job",
    subtitle = paste("Total Employee:", sum(proporsi_multiple$Jumlah))
  ) +
  theme_void() +
  theme(
    plot.title = element_text(size = 14, face = "bold", hjust = 0.5),
    plot.subtitle = element_text(size = 12, hjust = 0.5),
    legend.position = "bottom"
  ) +
  scale_fill_manual(values = c("Multiple Job" = "orange", "Single Job" = "steelblue"))

print(plot_pie)




# +++++++++++++++++++++++++++++++++++++++++++++++++++ #
#    Job dengan Overtime Salary Tertinggi (Case 3)
# +++++++++++++++++++++++++++++++++++++++++++++++++++ #

# Total Overtime salary per Job (3 tahun)
OT_perjob= MD_clean %>%
  group_by(`Job Code`, `Job`) %>%
  summarise(
    Total_Overtime = sum(Overtime, na.rm = TRUE),
    Rata2_Overtime = mean(Overtime, na.rm = TRUE),
    Jumlah_Pegawai = n_distinct(`Employee Identifier`),
    .groups = 'drop'
  ) %>%
  arrange(desc(Total_Overtime))
print(OT_perjob)

top10_OT <- OT_perjob %>%
  head(10)
View(top10_OT)

# Tren Overtime Salary per Job per Tahun
tren_overtime = MD_clean%>%
  group_by(`Job Code`, `Job`, Year) %>%
  summarise(
    Total_Overtime = sum(Overtime, na.rm = TRUE),
    Rata2_Overtime = mean(Overtime, na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  arrange(`Job Code`, Year)
print(head(tren_overtime, 20))

# Job dengan Overtime Salary Tertinggi per Tahun
top5_OTY = MD_clean %>%
  group_by(Year, `Job Code`, `Job`) %>%
  summarise(
    Total_Overtime = sum(Overtime, na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  group_by(Year) %>%
  arrange(Year, desc(Total_Overtime)) %>%
  slice_head(n = 5) %>%
  ungroup()
print(top5_OTY)
View(top5_OTY)

addWorksheet(wbgaji, sheetName = 'Top5_OTY')
writeData(wbgaji, sheet = "Top5_OTY", x=top5_OTY)

saveWorkbook(wbgaji, file = "Dataset_Portofolio1.xlsx", overwrite = TRUE)


# === Visualisasi Line Cart Case 3 === #
plot_tren = ggplot(top5_OTY, aes(x = factor(Year), 
                                 y = Total_Overtime/1000, 
                                 color = `Job` ,
                                 group = `Job`)) +
  geom_line(size = 1.5) +
  geom_point(size = 4) +
  geom_text(aes(label = round(Total_Overtime/1000, 2)), 
            vjust = -1, size = 3) +
  labs(
    title = "Tren Overtime Salary Top 5 Job (3 Tahun)",
    subtitle = "Overtime Salary per Tahun (USD)",
    x = "Tahun",
    y = "Total Overtime Salary (USD)",
    color = "Jabatan"
  ) +
  theme_minimal() +
  theme(
    legend.position = "bottom",
    plot.title = element_text(size = 14, face = "bold"),
    axis.text.x = element_text(size = 12)
  )
print(plot_tren)


# === Visualisasi Bar Cart Case 3 === #
data_items <- MD_clean %>%
  group_by(`Job Code`, `Job`, Year) %>%
  summarise(
    Total_Overtime = sum(Overtime, na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  group_by(`Job Code`) %>%
  mutate(Total_3Tahun = sum(Total_Overtime)) %>%
  ungroup() %>%
  arrange(desc(Total_3Tahun)) %>%
  distinct(`Job Code`, .keep_all = TRUE) %>%
  head(8) %>%
  mutate(
    Label = paste0(`Job Code`, "-" , Job),
  ) %>%
  select(`Job Code`, `Job`, Label) %>%
  left_join(
    MD_clean %>% 
      group_by(`Job Code`, `Job`, Year) %>%
      summarise(Total_Overtime = sum(Overtime, na.rm = TRUE), .groups = 'drop'),
    by = c("Job Code", "Job")
  )
View(data_items)

# Buat chart dengan label Items
plot_items <- ggplot(data_items, 
                     aes(x = reorder(Label, -Total_Overtime), 
                         y = Total_Overtime / 1000000,
                         fill = factor(Year))) +
  
  geom_col(position = position_dodge(width = 0.8), width = 0.7) +
  
  geom_text(aes(label = round(Total_Overtime / 1000000, 1)),
            position = position_dodge(width = 0.8),
            vjust = -0.5,
            size = 3,
            fontface = "bold") +
  
  scale_fill_brewer(palette = "Set2", name = "Tahun") +
  
  scale_y_continuous(
    limits = c(0, max(data_items$Total_Overtime / 1000000) * 1.15),
    breaks = seq(0, max(data_items$Total_Overtime / 1000000), by = 5),
    labels = paste0("$", seq(0, max(data_items$Total_Overtime / 1000000), by = 5), "M")
  ) +
  
  labs(
    title = "TOTAL OVERTIME SALARY PER JOB",
    subtitle = "Total Overtime Salary Tahun 2013-2015",
    x = "",
    y = "Total Overtime (USD)"
  ) +
  
  theme_minimal() +
  theme(
    plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
    plot.subtitle = element_text(size = 11, hjust = 0.5, color = "gray50"),
    
    axis.text.x = element_text(angle = 45, hjust = 1, size = 10, face = "bold"),
    axis.title.x = element_blank(),
    
    axis.text.y = element_text(size = 10, face = "bold"),
    axis.title.y = element_text(size = 11, face = "bold"),
    
    panel.grid.major.x = element_blank(),
    panel.grid.minor.y = element_blank(),
    panel.grid.major.y = element_line(color = "gray90", linewidth = 0.5),
    
    legend.position = "bottom",
    legend.title = element_text(size = 10, face = "bold")
  )

print(plot_items)
