
##### Sheet 1: 2025

```{r}

# Header aus Zeile 2 für Spaltennamen lesen
tmp_header <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                         sheet = 1, 
                         skip = 1,   # Header ist Zeile 2
                         n_max = 0) %>%
  names()

# Lücken im Header anpassen (i.e. erste Spalten benennen)
tmp_header <- c("Kanton", tmp_header)

tmp_name_map <- c("Kanton"        = "Kanton",
                  "Wahljahr 1"    = "Wahljahr",
                  "FDP 2"         = "FDP",
                  "SP"            = "SP",
                  "SVP"           = "SVP",
                  "LP 2"          = "LPS",
                  "EVP"           = "EVP",
                  "CSP"           = "CSP",
                  "GLP"           = "GLP",
                  "Die Mitte 3"   = "Mitte",
                  "CVP 3"         = "CVP",
                  "BDP 3"         = "BDP",
                  "PdA"           = "PdA",
                  "PSA"           = "PSA",
                  "Grüne 5"       = "Grüne",
                  "FGA"           = "FGA",
                  "Sol."          = "Sol.",
                  "Lega"          = "Lega",
                  "MCG (MCR)"     = "MCR",
                  "Übrige 4"      = "Übrige",
                  "Total"         = "Total")

# Automatisch ersetzen
tmp_header_neu <- ifelse(tmp_header %in% names(tmp_name_map),
                         tmp_name_map[tmp_header],
                         tmp_header)

print(tmp_header_neu)

# Daten ab Zeile 4 importieren
elec_kantonsregierung_raw <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                                        sheet = 1, 
                                        skip = 3,         # überspringt die ersten 3 Zeilen
                                        col_names = tmp_header_neu)

# Schritt 4: Nur Zeilen behalten, in denen "Wahljahr" nicht NA ist
elec_kantonsregierung_2025 <- elec_kantonsregierung_raw  %>%
  filter(!is.na(Wahljahr)) %>%
  select(Kanton, where(~ {x <- suppressWarnings(as.numeric(.x))
  sum(x, na.rm = TRUE) > 0}))

# Ergebnis anzeigen
print(elec_kantonsregierung_2025)
```

##### Sheet 2: 2024

```{r}


# Header aus Zeile 2 für Spaltennamen lesen
tmp_header <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                         sheet = 2, 
                         skip = 1,   # Header ist Zeile 2
                         n_max = 0) %>%
  names()

# Lücken im Header anpassen (i.e. erste Spalten benennen)
tmp_header <- c("Kanton", tmp_header)
cat(paste0('"', tmp_header, '"', ","), sep = "\n")

tmp_name_map <- c("Kanton"        = "Kanton",
                  "Wahljahr 1"    = "Wahljahr",
                  "FDP 2"         = "FDP",
                  "SP"            = "SP",
                  "SVP"           = "SVP",
                  "LP 2"          = "LPS",
                  "EVP"           = "EVP",
                  "CSP"           = "CSP",
                  "GLP"           = "GLP",
                  "Die Mitte 3"   = "Mitte",
                  "CVP 3"         = "CVP",
                  "BDP 3"         = "BDP",
                  "PdA"           = "PdA",
                  "PSA"           = "PSA",
                  "Grüne 5"       = "Grüne",
                  "FGA"           = "FGA",
                  "Sol."          = "Sol.",
                  "Lega"          = "Lega",
                  "MCG (MCR)"     = "MCR",
                  "Übrige 4"      = "Übrige",
                  "Total"         = "Total")

# Automatisch ersetzen
tmp_header_neu <- ifelse(tmp_header %in% names(tmp_name_map),
                         tmp_name_map[tmp_header],
                         tmp_header)

print(tmp_header_neu)

# Daten ab Zeile 4 importieren
elec_kantonsregierung_raw <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                                        sheet = 2, 
                                        skip = 3,         # überspringt die ersten 3 Zeilen
                                        col_names = tmp_header_neu)

# Schritt 4: Nur Zeilen behalten, in denen "Wahljahr" nicht NA ist
elec_kantonsregierung_2024 <- elec_kantonsregierung_raw  %>%
  filter(!is.na(Wahljahr)) %>%
  select(Kanton, where(~ {x <- suppressWarnings(as.numeric(.x))
  sum(x, na.rm = TRUE) > 0}))

# Ergebnis anzeigen
print(elec_kantonsregierung_2024)
```

##### Sheet 3: 2023

```{r}


# Header aus Zeile 2 für Spaltennamen lesen
tmp_header <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                         sheet = 3, 
                         skip = 1,   # Header ist Zeile 2
                         n_max = 0) %>%
  names()

# Lücken im Header anpassen (i.e. erste Spalten benennen)
tmp_header <- c("Kanton", tmp_header)
cat(paste0('"', tmp_header, '"', ","), sep = "\n")

tmp_name_map <- c("Kanton"        = "Kanton",
                  "Wahljahr 1)"    = "Wahljahr",
                  "FDP 2)"         = "FDP",
                  "SP"            = "SP",
                  "SVP"           = "SVP",
                  "LP 2)"          = "LPS",
                  "EVP"           = "EVP",
                  "CSP"           = "CSP",
                  "GLP"           = "GLP",
                  "Die Mitte 3)"   = "Mitte",
                  "CVP 3)"         = "CVP",
                  "BDP 3)"         = "BDP",
                  "PdA"           = "PdA",
                  "PSA"           = "PSA",
                  "Grüne 5)"       = "Grüne",
                  "FGA"           = "FGA",
                  "Sol."          = "Sol.",
                  "Lega"          = "Lega",
                  "MCG (MCR)"     = "MCR",
                  "Übrige 4)"      = "Übrige",
                  "Total"         = "Total")

# Automatisch ersetzen
tmp_header_neu <- ifelse(tmp_header %in% names(tmp_name_map),
                         tmp_name_map[tmp_header],
                         tmp_header)

print(tmp_header_neu)

# Daten ab Zeile 4 importieren
elec_kantonsregierung_raw <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                                        sheet = 3, 
                                        skip = 3,         # überspringt die ersten 3 Zeilen
                                        col_names = tmp_header_neu)

# Schritt 4: Nur Zeilen behalten, in denen "Wahljahr" nicht NA ist
elec_kantonsregierung_2023 <- elec_kantonsregierung_raw  %>%
  filter(!is.na(Wahljahr)) %>%
  select(Kanton, where(~ {x <- suppressWarnings(as.numeric(.x))
  sum(x, na.rm = TRUE) > 0}))

# Ergebnis anzeigen
print(elec_kantonsregierung_2023)
```

##### Sheet 4: 2022

```{r}


# Header aus Zeile 2 für Spaltennamen lesen
tmp_header <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                         sheet = 4, 
                         skip = 1,   # Header ist Zeile 2
                         n_max = 0) %>%
  names()

# Lücken im Header anpassen (i.e. erste Spalten benennen)
tmp_header <- c("Kanton", tmp_header)
cat(paste0('"', tmp_header, '"', ","), sep = "\n")

tmp_name_map <- c("Kanton"        = "Kanton",
                  "Wahljahr 1)"    = "Wahljahr",
                  "FDP 2)"         = "FDP",
                  "SP"            = "SP",
                  "SVP"           = "SVP",
                  "LP 2)"          = "LPS",
                  "EVP"           = "EVP",
                  "CSP"           = "CSP",
                  "GLP"           = "GLP",
                  "Die Mitte 3)"   = "Mitte",
                  "CVP 3)"         = "CVP",
                  "BDP 3)"         = "BDP",
                  "PdA"           = "PdA",
                  "PSA"           = "PSA",
                  "Grüne 6)"       = "Grüne",
                  "FGA"           = "FGA",
                  "Sol."          = "Sol.",
                  "Lega"          = "Lega",
                  "MCR"     = "MCR",
                  "Übrige 4)"      = "Übrige",
                  "Total"         = "Total")

# Automatisch ersetzen
tmp_header_neu <- ifelse(tmp_header %in% names(tmp_name_map),
                         tmp_name_map[tmp_header],
                         tmp_header)

print(tmp_header_neu)

# Daten ab Zeile 4 importieren
elec_kantonsregierung_raw <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                                        sheet = 4, 
                                        skip = 3,         # überspringt die ersten 3 Zeilen
                                        col_names = tmp_header_neu)

# Schritt 4: Nur Zeilen behalten, in denen "Wahljahr" nicht NA ist
elec_kantonsregierung_2022 <- elec_kantonsregierung_raw  %>%
  filter(!is.na(Wahljahr)) %>%
  select(Kanton, where(~ {x <- suppressWarnings(as.numeric(.x))
  sum(x, na.rm = TRUE) > 0}))

# Ergebnis anzeigen
print(elec_kantonsregierung_2022)

```

##### Sheet 5: 2021

```{r}


# Header aus Zeile 2 für Spaltennamen lesen
tmp_header <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                         sheet = 5, 
                         skip = 1,   # Header ist Zeile 2
                         n_max = 0) %>%
  names()

# Lücken im Header anpassen (i.e. erste Spalten benennen)
tmp_header <- c("Kanton", tmp_header)
cat(paste0('"', tmp_header, '"', ","), sep = "\n")

tmp_name_map <- c("Kanton"        = "Kanton",
                  "Wahljahr 1)"    = "Wahljahr",
                  "FDP 2)"         = "FDP",
                  "SP"            = "SP",
                  "SVP"           = "SVP",
                  "LP 2)"          = "LPS",
                  "EVP"           = "EVP",
                  "CSP"           = "CSP",
                  "GLP"           = "GLP",
                  "Die Mitte 3)"   = "Mitte",
                  "CVP 3)"         = "CVP",
                  "BDP 5)"         = "BDP",
                  "PdA"           = "PdA",
                  "PSA"           = "PSA",
                  "GPS"       = "Grüne",
                  "FGA"           = "FGA",
                  "Sol."          = "Sol.",
                  "Lega"          = "Lega",
                  "MCR"     = "MCR",
                  "Übrige 4)"      = "Übrige",
                  "Total"         = "Total")

# Automatisch ersetzen
tmp_header_neu <- ifelse(tmp_header %in% names(tmp_name_map),
                         tmp_name_map[tmp_header],
                         tmp_header)

print(tmp_header_neu)

# Daten ab Zeile 4 importieren
elec_kantonsregierung_raw <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                                        sheet = 5, 
                                        skip = 3,         # überspringt die ersten 3 Zeilen
                                        col_names = tmp_header_neu)

# Schritt 4: Nur Zeilen behalten, in denen "Wahljahr" nicht NA ist
elec_kantonsregierung_2021 <- elec_kantonsregierung_raw  %>%
  filter(!is.na(Wahljahr)) %>%
  select(Kanton, where(~ {x <- suppressWarnings(as.numeric(.x))
  sum(x, na.rm = TRUE) > 0}))

# Ergebnis anzeigen
print(elec_kantonsregierung_2021)

```

##### Sheet 6: 2020

```{r}

# Header aus Zeile 2 für Spaltennamen lesen
tmp_header <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                         sheet = 6, 
                         skip = 1,   # Header ist Zeile 2
                         n_max = 0) %>%
  names()

# Lücken im Header anpassen (i.e. erste Spalten benennen)
tmp_header <- c("Kanton", tmp_header)
cat(paste0('"', tmp_header, '"', ","), sep = "\n")

tmp_name_map <- c("Kanton"         = "Kanton",
                  "Wahljahr 1)"    = "Wahljahr",
                  "FDP 2)"         = "FDP",
                  "CVP 3)"         = "CVP",
                  "SP"             = "SP",
                  "SVP"            = "SVP",
                  "LP 2)"          = "LPS",
                  "CSP"            = "CSP",
                  "GLP"            = "GLP",
                  "BDP"            = "BDP",
                  "GPS"            = "Grüne",
                  "Lega"           = "Lega",
                  "MCR"            = "MCR",
                  "Übrige 4)"      = "Übrige",
                  "Total"          = "Total")

# Automatisch ersetzen
tmp_header_neu <- ifelse(tmp_header %in% names(tmp_name_map),
                         tmp_name_map[tmp_header],
                         tmp_header)

print(tmp_header_neu)

# Daten ab Zeile 4 importieren
elec_kantonsregierung_raw <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                                        sheet = 6, 
                                        skip = 3,         # überspringt die ersten 3 Zeilen
                                        col_names = tmp_header_neu)

# Schritt 4: Nur Zeilen behalten, in denen "Wahljahr" nicht NA ist
elec_kantonsregierung_2020 <- elec_kantonsregierung_raw  %>%
  filter(!is.na(Wahljahr)) %>%
  select(Kanton, where(~ {x <- suppressWarnings(as.numeric(.x))
  sum(x, na.rm = TRUE) > 0}))


# Ergebnis anzeigen
print(elec_kantonsregierung_2020)

```

##### Sheet 7: 2019

```{r}


# Header aus Zeile 2 für Spaltennamen lesen
tmp_header <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                         sheet = 7, 
                         skip = 1,   # Header ist Zeile 2
                         n_max = 0) %>%
  names()

# Lücken im Header anpassen (i.e. erste Spalten benennen)
tmp_header <- c("Kanton", tmp_header)
cat(paste0('"', tmp_header, '"', ","), sep = "\n")

tmp_name_map <- c("Kanton"        = "Kanton",
                  "Wahljahr 1)"    = "Wahljahr",
                  "FDP 2)"         = "FDP",
                  "SP"            = "SP",
                  "SVP"           = "SVP",
                  "LP 2)"          = "LPS",
                  "EVP"           = "EVP",
                  "CSP"           = "CSP",
                  "GLP"           = "GLP",
                  "Die Mitte 3)"   = "Mitte",
                  "CVP 3)"         = "CVP",
                  "BDP 3)"         = "BDP",
                  "PdA"           = "PdA",
                  "PSA"           = "PSA",
                  "Grüne 5)"       = "Grüne",
                  "FGA"           = "FGA",
                  "Sol."          = "Sol.",
                  "Lega"          = "Lega",
                  "MCG (MCR)"     = "MCR",
                  "Übrige 4)"      = "Übrige",
                  "Total"         = "Total")

# Automatisch ersetzen
tmp_header_neu <- ifelse(tmp_header %in% names(tmp_name_map),
                         tmp_name_map[tmp_header],
                         tmp_header)

print(tmp_header_neu)

# Daten ab Zeile 4 importieren
elec_kantonsregierung_raw <- read_excel("~/CAS/Zertifikatsarbeit/data/elections/je-d-17.02.06.01_KANTON_Kantonale_Regierungswahlen.xlsx", 
                                        sheet = 7, 
                                        skip = 3,         # überspringt die ersten 3 Zeilen
                                        col_names = tmp_header_neu)

# Schritt 4: Nur Zeilen behalten, in denen "Wahljahr" nicht NA ist
elec_kantonsregierung_2019 <- elec_kantonsregierung_raw  %>%
  filter(!is.na(Wahljahr)) %>%
  select(Kanton, where(~ {x <- suppressWarnings(as.numeric(.x))
  sum(x, na.rm = TRUE) > 0}))

# Ergebnis anzeigen
print(elec_kantonsregierung_2019)
```

##### Sheets konsolidieren & Change log erstellen

```{r}

# Alle Dataframes in Liste packen
df_list <- list(elec_kantonsregierung_2019,
                elec_kantonsregierung_2020,
                elec_kantonsregierung_2021,
                elec_kantonsregierung_2022,
                elec_kantonsregierung_2023,
                elec_kantonsregierung_2024,
                elec_kantonsregierung_2025)

# Alle Dataframes zusammenführen (Vereinigung aller Spalten)
df_all <- bind_rows(df_list)

# Filter auf relevante Wahljahre setzen (2015, da Legislatur 4-5 Jahre dauert und später der Zeitraum ab 2019 geprüft wird)
df_all <- df_all %>%
  filter(Wahljahr >= 2015, Wahljahr <= 2025)

# Hilfsfuntion um alle Einträge in numerische umzuwandeln
make_numeric <- function(x) suppressWarnings(as.numeric(x))

# Funktion auf alle Spalten ausser Kanton & Wahljahr anwenden.
df_all <- df_all %>%
  mutate(across(-c(Kanton, Wahljahr), make_numeric))

# Pro Kanton+Wahljahr auf eine Zeile reduzieren
df_all <- df_all %>%
  group_by(Kanton, Wahljahr) %>%
  summarise(across(everything(), ~ coalesce(last(na.omit(.)), NA_real_)), .groups = "drop")

# Alle Kantone und Wahljahre eruieren/bestimmen, die in den Daten vorkommen
vektor_kanton <- unique(df_all$Kanton)
vektor_wahljahr <- 2015:2025

# Vollständigen Grid erzeugen
df_grid <- expand_grid(Kanton = vektor_kanton,
                       Wahljahr = vektor_wahljahr)

# Mit den Daten verknüpfen
df_all <- left_join(df_grid,
                    df_all,
                    by = c("Kanton", "Wahljahr"))


elec_kantonsregierung_final <- df_all %>%
  arrange(Kanton, Wahljahr) %>%
  group_by(Kanton) %>%
  fill(everything(), .direction = "down") %>%
  ungroup()


```