#------ Librerías --------#
library(readxl)
library(dplyr)
library(openxlsx)
library(tidyr)
#-------------------------#

#------ Rutas --------#
ruta_postulados <- "C:/Users/jorge.romero.c/Downloads/Postulados.xlsx"
ruta_reporte <- "C:/Users/jorge.romero.c/Downloads/Reporte de Producción Académica (18-07-2025).xlsx"
ruta_matriz <- "C:/Users/jorge.romero.c/Downloads/PropuestaMatrizDeProductosAcademicos2025.xlsx"
#---------------------#

#------ Lectura de archivos --------#
postulaciones <- read_excel(ruta_postulados, sheet = "Datos") %>%
  mutate(`ID de persona` = as.character(`ID de persona`))

reporte <- read_excel(ruta_reporte, sheet = "FORMACIÓN DE PRODUCCIÓN ACADÉMI") %>%
  mutate(`ID de sistema de usuario` = as.character(`ID de sistema de usuario`))

habilitantes <- read_excel(ruta_matriz, sheet = "ProducciónAcadémicaEscalafón25")
cat("Productos totales cargados:", nrow(reporte), "\n")
#----------------------------------#

#------1. Filtrar solo profesores que se postularon ------##Verificado#
reporte <- reporte %>%
  left_join(
    postulaciones %>%
      select(`ID de persona`, `Escalafón a la cual me postulo (Picklist Label)`),
    by = c("ID de sistema de usuario" = "ID de persona")
  ) %>%
  filter(!is.na(`Escalafón a la cual me postulo (Picklist Label)`))

cat("Productos tras filtrar profesores postulados:", nrow(reporte), "\n")
cat("Profesores postulados con productos:", n_distinct(reporte$`ID de sistema de usuario`), "\n") 

#------2. Filtrar por Usabilidad = Escalafón ------#
reporte <- reporte %>% filter(Usabilidad == "Escalafón")
cat("Productos tras filtrar Usabilidad 'Escalafón':", nrow(reporte), "\n")
cat("Profesores postulados con productos:", n_distinct(reporte$`ID de sistema de usuario`), "\n") 

#------3. Eliminar productos con nombre y sub_subcategoría vacíos ------#
reporte <- reporte %>%
  filter(!(is.na(`Nombre de la Producción`) & is.na(Sub_Subcategoria)) &
           !(`Nombre de la Producción` == "" & Sub_Subcategoria == ""))
cat("Productos tras eliminar vacíos en nombre y sub_subcategoría:", nrow(reporte), "\n")
cat("Profesores postulados con productos:", n_distinct(reporte$`ID de sistema de usuario`), "\n") 

#------ Agregar Escalafón postulado ------#
reporte <- reporte %>%
  left_join(
    postulaciones %>%
      select(`ID de persona`, EscalafonPostulado = `Escalafón a la cual me postulo (Picklist Label)`),
    by = c("ID de sistema de usuario" = "ID de persona")
  ) %>%
  mutate(EscalafonPostulado = paste("Habilitantes", EscalafonPostulado))

cat("Productos tras unir con escalafón postulado:", nrow(reporte), "\n")
cat("Profesores postulados con productos:", n_distinct(reporte$`ID de sistema de usuario`), "\n") 

#------ Transformar matriz habilitantes a formato largo ------#
habilitantes_largos <- habilitantes %>%
  select(`Categoría y tipo de producto`,
         `Habilitantes Asistente 2`,
         `Habilitantes Asociado 1`,
         `Habilitantes Asociado 2`,
         `Habilitantes Titular`) %>%
  pivot_longer(
    cols = starts_with("Habilitantes"),
    names_to = "Escalafon",
    values_to = "Habilitante"
  ) %>%
  filter(trimws(Habilitante) == "Habilitante") %>%
  select(`Categoría y tipo de producto`, Escalafon, Habilitante)

cat("Total combinaciones de productos habilitantes:", nrow(habilitantes_largos), "\n")

#------ Comparar productos contra matriz de habilitantes ------#
reporte_validado <- reporte %>%
  left_join(habilitantes_largos,
            by = c("Sub_Subcategoria" = "Categoría y tipo de producto",
                   "EscalafonPostulado" = "Escalafon"))
cat("Productos tras validación contra matriz:", nrow(reporte_validado), "\n")
cat("Profesores postulados con productos:", n_distinct(reporte_validado$`ID de sistema de usuario`), "\n") 

#------ Definir si requiere validación y si es habilitante ------#
escalafones_sin_validacion <- c(
  "Habilitantes Instructor 1",
  "Habilitantes Instructor 2",
  "Habilitantes Asistente 1"
)

reporte_validado <- reporte_validado %>%
  mutate(
    requiere_validacion = !(EscalafonPostulado %in% escalafones_sin_validacion),
    es_habilitante = case_when(
      requiere_validacion ~ !is.na(Habilitante),
      !requiere_validacion & !is.na(Sub_Subcategoria) ~ TRUE,
      TRUE ~ FALSE
    )
  )

cat("Productos marcados como habilitantes:", sum(reporte_validado$es_habilitante), "\n")
cat("Productos NO habilitantes:", sum(!reporte_validado$es_habilitante), "\n")

# Paso 1: Profesores que NO requieren validación (escalafones que no exigen habilitante)
profesores_sin_validacion <- reporte_validado %>%
  filter(EscalafonPostulado %in% escalafones_sin_validacion) %>%
  distinct(`ID de sistema de usuario`)
cat("Productos tras validación contra matriz:", nrow(profesores_sin_validacion), "\n")

# Paso 2: Profesores que SÍ requieren habilitante y tienen al menos uno válido
profesores_con_habilitante <- reporte_validado %>%
  filter(requiere_validacion, es_habilitante) %>%
  distinct(`ID de sistema de usuario`)
cat("Productos tras validación contra matriz:", nrow(profesores_con_habilitante), "\n")

# Paso 3: Unión de ambos conjuntos de profesores válidos
profes_validos <- bind_rows(profesores_sin_validacion, profesores_con_habilitante) %>%
  distinct(`ID de sistema de usuario`)

cat("Profesores válidos para calificación:", nrow(profes_validos), "\n")

# Paso 4: Traer todos los productos de esos profesores
reporte_final <- reporte_validado %>%
  filter(`ID de sistema de usuario` %in% profes_validos$`ID de sistema de usuario`)

cat("Total productos finales a calificar:", nrow(reporte_final), "\n")

# Especifica la ruta de salida
ruta_salida <- "C:/Users/jorge.romero.c/Downloads/Productos_Validados.xlsx"

# Guarda el data frame en un archivo Excel
openxlsx::write.xlsx(reporte_final, file = ruta_salida)

cat("✅ Archivo exportado exitosamente en:", ruta_salida, "\n")

# Productos que SÍ deben ser validados
reporte_validados_finales <- reporte_validado %>%
  filter(`ID de sistema de usuario` %in% profes_validos$`ID de sistema de usuario`)

# Productos EXCLUIDOS con motivo
reporte_excluidos <- reporte_validado %>%
  filter(!(`ID de sistema de usuario` %in% profes_validos$`ID de sistema de usuario`)) %>%
  mutate(
    motivo_exclusion = case_when(
      EscalafonPostulado %in% escalafones_sin_validacion & is.na(Sub_Subcategoria) ~ "Sin subcategoría válida",
      requiere_validacion & is.na(Habilitante) ~ "Sin producto habilitante válido para su escalafón",
      TRUE ~ "Otro"
    )
  )

cat("Productos a validar:", nrow(reporte_validados_finales), "\n")
cat("Productos excluidos:", nrow(reporte_excluidos), "\n")

# Exportar ambos DataFrames
ruta_validados <- "C:/Users/jorge.romero.c/Downloads/Productos_Validados2.xlsx"
ruta_excluidos <- "C:/Users/jorge.romero.c/Downloads/Productos_Excluidos.xlsx"

write.xlsx(reporte_validados_finales, file = ruta_validados)
write.xlsx(reporte_excluidos, file = ruta_excluidos)

cat("✅ Archivos exportados:\n")
cat("✔️ Validados:", ruta_validados, "\n")
cat("✔️ Excluidos:", ruta_excluidos, "\n")

#------ Preparar y exportar archivos por sede (formato nuevo) --------#
sedes <- unique(reporte_validado$Sede)
ruta_base <- "C:/Users/jorge.romero.c/Downloads/Reporte_Escalafon_"

for (sede in sedes) {
  # 1. Todos los productos válidos de esta sede
  productos_validos <- reporte_validado %>%
    filter(Sede == sede,
           `ID de sistema de usuario` %in% profes_validos$`ID de sistema de usuario`)
  
  # 2. Todos los productos excluidos de esta sede
  productos_excluidos <- reporte_validado %>%
    filter(Sede == sede,
           !(`ID de sistema de usuario` %in% profes_validos$`ID de sistema de usuario`)) %>%
    mutate(
      motivo_exclusion = case_when(
        EscalafonPostulado %in% escalafones_sin_validacion & is.na(Sub_Subcategoria) ~ "Sin subcategoría válida",
        requiere_validacion & is.na(Habilitante) ~ "Sin producto habilitante válido para su escalafón",
        TRUE ~ "Otro"
      )
    )
  
  # 3. De los productos válidos: los de escalafones que NO requieren habilitantes
  sin_habilitante_requerido <- productos_validos %>%
    filter(!requiere_validacion)
  
  # 4. De los productos válidos: los de escalafones que SÍ requieren habilitantes y tienen al menos uno válido
  con_habilitante_requerido <- productos_validos %>%
    filter(requiere_validacion)
  
  # Exportar archivo con las 4 hojas para esta sede
  openxlsx::write.xlsx(
    list(
      Productos_Validos = productos_validos,
      Productos_Excluidos = productos_excluidos,
      Sin_Habilitante_Requerido = sin_habilitante_requerido,
      Con_Habilitante_Requerido = con_habilitante_requerido
    ),
    file = paste0(ruta_base, gsub("[^A-Za-z0-9]", "_", sede), ".xlsx")
  )
}

cat("✅ Archivos por sede generados exitosamente con nuevo formato.\n")


###############################################################################
###############################################################################


