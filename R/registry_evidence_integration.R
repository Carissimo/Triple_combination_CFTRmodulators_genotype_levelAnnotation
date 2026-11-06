###############################################################################
# REGISTRY-BASED EVIDENCE-INTEGRATED FRAMEWORK
#
# WHAT THIS SCRIPT DOES
# Starting from a registry of patients (one row per patient, two CFTR variants
# each) it:
#   1. Classifies every patient on eligibility criteria (two true Class I
#      variants, at least one F508del, at least one variant designated
#      responsive by EMA for ETI or VTD) and saves the annotated patient table.
#   2. Keeps only the patients that meet NONE of those criteria and builds the
#      "evidence-integrated framework" table, adding for each of their two
#      variants the evidence available from the French cohort (real-world
#      data), the FDA lists and the in vitro studies.
#
# Variants are identified by their LEGACY NAME (e.g. "G542X", "F508del",
# "R74W;V201M;D1270N"). Complex alleles are written as components separated
# by ";".
#
# REQUIRED PACKAGES
#   install.packages(c("readxl", "openxlsx"))
#
# QUICK START (a single call does everything)
#   source("registry_evidence_integration.R")
#
#   files <- list(
#     patients           = "PATIENTS_FILE.xlsx",
#     framework_template = "Registry_based_evidence_integrated_framework_output.xlsx",
#     true_class1        = "ETI_real_world_data_true_class__I_FCP.xlsx",
#     exceptional_c1     = "ETI_real_world_data_exceptional_class__I_FCP.xlsx",
#     fcp_rwd            = "ETI_real_world_data_FCP.xlsx",
#     ema_eti            = "ETI_responsive_variants_EMA.xlsx",
#     ema_vtd            = "VTD_responsive_variants_EMA.xlsx",
#     fda_eti            = "Variants_individually_listed_as_responsive_to_ETI_FDA.xlsx",
#     fda_vtd            = "Variants_individually_listed_as_responsive_to_VTD_FDA.xlsx",
#     in_vitro_eti       = "655_variants_tested_in_vitro_with_ETI.xlsx",
#     in_vitro_vti       = "417_variants_tested_in_vitro_with_VTI.xlsx"
#   )
#   res <- run_registry_framework(files, output_dir = "output")
#
#   # If all files are in one folder and have the names above:
#   res <- run_registry_framework(default_file_list("path/to/folder"))
#
#   # Same, with a patients file that has a different name:
#   res <- run_registry_framework(default_file_list("path/to/folder",
#                                                   patients_file = "my_registry.xlsx"))
#
# INPUT FILES (all .xlsx, first sheet)
#   patients           Column 1 = patient ID, column 2 = variant 1,
#                      column 3 = variant 2 (legacy names). Any column titles.
#   framework_template Empty table (header row only) that defines the columns
#                      of the framework output. See "FRAMEWORK TEMPLATE" below.
#   true_class1        List of true Class I variants (legacy names in column 1).
#   exceptional_c1     List of exceptional Class I variants (column 1).
#   fcp_rwd            French cohort real-world data: one column per category
#                      (Responsive, Probably responsive, Probably
#                      non-responsive, Non-responsive, Inconclusive).
#   ema_eti, ema_vtd   Variants designated responsive by EMA for ETI / VTD
#                      (column 1).
#   fda_eti, fda_vtd   Variants listed as responsive by FDA for ETI / VTD
#                      (column 1).
#   in_vitro_eti,      In vitro results for ETI / VTD: one column per category
#   in_vitro_vti       (Responsive, Borderline responsive, Non-responsive).
#
# OUTPUT FILES (written to output_dir)
#   <patients file name>_annotated.xlsx
#       The patients table plus the eligibility columns.
#   <framework template file name>.xlsx
#       The framework template filled with the selected patients and the
#       evidence columns.
###############################################################################

suppressPackageStartupMessages({
  library(readxl)     # reading .xlsx files
  library(openxlsx)   # writing .xlsx files
})

# ----------------------------------------------------------------------------
# HELPER FUNCTIONS
# ----------------------------------------------------------------------------

# Cleans a variant string so that formatting differences do not prevent a
# match: converts non-breaking spaces to normal spaces, removes spaces around
# ";", collapses repeated spaces and trims the ends.
#   "F508C;S1251N "   -> "F508C;S1251N"
#   "711+1G->T; S977F" -> "711+1G->T;S977F"
# Empty strings become NA.
norm_variant <- function(x) {
  x <- as.character(x)
  x <- gsub("\u00a0", " ", x, fixed = TRUE)
  x <- gsub("\\s*;\\s*", ";", x)
  x <- gsub("\\s+", " ", x)
  x <- trimws(x)
  x[!is.na(x) & x == ""] <- NA_character_
  x
}

# Builds the key used to compare variants. The variant is cleaned with
# norm_variant() and, if ignore_order = TRUE, the components of a complex
# allele are sorted alphabetically, so that "A1006E;V562I" and "V562I;A1006E"
# (or "D443Y;G576A;R668C" and "R668C;D443Y;G576A") produce the same key.
# Keys are used only for matching: the values written in the output files are
# the original ones.
match_key <- function(x, ignore_order = TRUE) {
  x <- norm_variant(x)
  if (!ignore_order) return(x)
  vapply(x, function(s) {
    if (is.na(s)) return(NA_character_)
    paste(sort(strsplit(s, ";", fixed = TRUE)[[1]], method = "radix"),
          collapse = ";")
  }, character(1), USE.NAMES = FALSE)
}

# Reads the first column of a file and returns the set of its variants
# (as matching keys, without empty cells or duplicates).
read_first_col <- function(path, sheet = 1, ignore_order = TRUE) {
  d <- read_excel(path, sheet = sheet, col_types = "text")
  v <- match_key(d[[1]], ignore_order)
  unique(v[!is.na(v)])
}

# TRUE for each element of x that is present in set. NA is never a match.
# The whole allele must be equal to an entry of the list.
in_set <- function(x, set) !is.na(x) & x %in% set

# TRUE for each allele that contains F508del as one of its components, also
# inside a complex allele ("F508del;5T;TG12", "F508del;621+1G>T").
# The allele is split at ";" and each component must be exactly "F508del":
# "F508C" or "F508S" are different variants and do not count.
has_f508del <- function(x) {
  vapply(x, function(s) {
    if (is.na(s)) return(FALSE)
    "F508del" %in% trimws(strsplit(s, ";", fixed = TRUE)[[1]])
  }, logical(1), USE.NAMES = FALSE)
}

# Reads a file whose columns are categories (column title = category name,
# cells = variants). Returns a named vector: names are the variants (keys),
# values are the category (column title). If a variant is listed under more
# than one category a warning is issued and the first category is kept.
read_category_lookup <- function(path, sheet = 1, ignore_order = TRUE) {
  d <- read_excel(path, sheet = sheet, col_types = "text")
  out <- character(0)
  for (cat in names(d)) {
    v <- match_key(d[[cat]], ignore_order); v <- unique(v[!is.na(v)])
    dup <- intersect(v, names(out))
    if (length(dup))
      warning("Variants listed in more than one category (", basename(path), "): ",
              paste(dup, collapse = ", "), call. = FALSE)
    v <- setdiff(v, names(out))
    out <- c(out, setNames(rep(cat, length(v)), v))
  }
  out
}

# Reads an in vitro file (ETI or VTI). Its columns are titled "Responsive
# variants", "Borderline responsive variants" and "Non-responsive variants";
# the category labels returned are "Responsive", "Borderline" and
# "Non-responsive".
read_in_vitro_lookup <- function(path, ignore_order = TRUE) {
  iv <- read_category_lookup(path, ignore_order = ignore_order)
  lab <- sub(" variants$", "", unname(iv), perl = TRUE)
  lab <- sub("^Borderline responsive$", "Borderline", lab)
  setNames(lab, names(iv))
}

# Looks up the category of each variant in x using a named vector produced by
# read_category_lookup(). Variants not found (or NA) get NA.
lookup_category <- function(x, lookup) {
  res <- unname(lookup[x])
  res[is.na(x)] <- NA_character_
  res
}

# Converts TRUE/FALSE into "X"/empty and "yes"/"no".
mark_x <- function(flag) ifelse(flag, "X", NA_character_)
yes_no <- function(flag) ifelse(flag, "yes", "no")

# FRAMEWORK TEMPLATE
# The columns of the framework output are defined by the header row of the
# template file, which can be edited freely (titles, order, extra columns).
# Each column is recognized by key words in its title, not by its position:
#   Patient ID                      "Patient ID"
#   Variant legacy names            "VARIANT 1" / "VARIANT 2" + "legacy"
#   French cohort RWD               "VARIANT 1" / "VARIANT 2" + "RWD"
#   FDA ETI / FDA VTD               "VARIANT n" + "FDA" + "ETI" (or "VTD")
#   In vitro ETI                    "VARIANT n" + "in vitro" + "ETI"
#   In vitro VTI                    "VARIANT n" + "in vitro" without "ETI"
# Key words are not case-sensitive and can appear in any order in the title.
# Columns of the template that are not recognized are left empty.
#
# Returns the position of the single header that contains all the given
# regular expressions; stops with an explanatory error if there is none or
# more than one. 'what' is the description used in the error message.
find_template_col <- function(headers, ..., what) {
  pats <- c(...)
  hit <- rep(TRUE, length(headers))
  for (pt in pats) hit <- hit & grepl(pt, trimws(headers), ignore.case = TRUE, perl = TRUE)
  idx <- which(hit)
  if (length(idx) != 1)
    stop("FRAMEWORK template: expected exactly one column for '", what,
         "' but found ", length(idx), ". Headers found:\n  ",
         paste(headers, collapse = "\n  "), call. = FALSE)
  idx
}

# Names that the list of input files must contain (see QUICK START).
REQUIRED_FILES <- c("patients", "framework_template", "true_class1",
                    "exceptional_c1", "fcp_rwd", "ema_eti", "ema_vtd",
                    "fda_eti", "fda_vtd", "in_vitro_eti", "in_vitro_vti")

# Builds the list of input files for run_registry_framework() when all of them
# are in the same folder (input_dir) and have the standard names. The patients
# file name can be changed with patients_file.
default_file_list <- function(input_dir = ".", patients_file = "PATIENTS_FILE.xlsx") {
  f <- list(
    patients       = patients_file,
    framework_template = "Registry_based_evidence_integrated_framework_output.xlsx",
    true_class1    = "ETI_real_world_data_true_class__I_FCP.xlsx",
    exceptional_c1 = "ETI_real_world_data_exceptional_class__I_FCP.xlsx",
    fcp_rwd        = "ETI_real_world_data_FCP.xlsx",
    ema_eti        = "ETI_responsive_variants_EMA.xlsx",
    ema_vtd        = "VTD_responsive_variants_EMA.xlsx",
    fda_eti        = "Variants_individually_listed_as_responsive_to_ETI_FDA.xlsx",
    fda_vtd        = "Variants_individually_listed_as_responsive_to_VTD_FDA.xlsx",
    in_vitro_eti   = "655_variants_tested_in_vitro_with_ETI.xlsx",
    in_vitro_vti   = "417_variants_tested_in_vitro_with_VTI.xlsx"
  )
  lapply(f, function(x) file.path(input_dir, x))
}

# ----------------------------------------------------------------------------
# MAIN FUNCTION
# ----------------------------------------------------------------------------
# Arguments
#   files        Named list of input file paths (see QUICK START); it must
#                contain the names listed in REQUIRED_FILES.
#   output_dir   Folder where the two output files are written (created if it
#                does not exist).
#   write_files  If FALSE nothing is written to disk and the results are only
#                returned.
#   ignore_order If TRUE (default) the order of the components of a complex
#                allele does not matter when matching ("A;B" equals "B;A").
#                If FALSE the allele must be written in the same order as in
#                the lists.
#
# Value (returned invisibly): a list with
#   $patients   the annotated patients table (all patients)
#   $framework  the framework table (selected patients with evidence columns)
run_registry_framework <- function(files,
                                   output_dir   = "output",
                                   write_files  = TRUE,
                                   ignore_order = TRUE) {

  # --- Check that all the input files are listed and exist -------------------
  files <- as.list(files)
  missing_names <- setdiff(REQUIRED_FILES, names(files))
  if (length(missing_names))
    stop("Missing entries in 'files': ", paste(missing_names, collapse = ", "),
         call. = FALSE)
  not_found <- unlist(files[REQUIRED_FILES])
  not_found <- not_found[!file.exists(not_found)]
  if (length(not_found))
    stop("File(s) not found:\n  ", paste(not_found, collapse = "\n  "), call. = FALSE)
  if (write_files) dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

  # --------------------------------------------------------------------------
  # STEP 1 - Read the patients file
  # --------------------------------------------------------------------------
  # Only the first three columns are used (ID, variant 1, variant 2), whatever
  # their titles; they are given standard titles in the output.
  patients <- as.data.frame(read_excel(files$patients, col_types = "text"),
                            check.names = FALSE)
  patients <- patients[, 1:3]
  names(patients) <- c("Patient ID",
                       "Variant 1 (Legacy name)",
                       "Variant 2 (Legacy name)")
  patients <- patients[!is.na(patients[[1]]), ]     # drop rows without an ID
  patients[[1]] <- as.integer(patients[[1]])
  stopifnot(!anyDuplicated(patients[[1]]))          # each patient ID must be unique

  # Matching keys of the two variants of every patient (the original values
  # in 'patients' are left untouched and are the ones written to the output).
  v1 <- match_key(patients[[2]], ignore_order)
  v2 <- match_key(patients[[3]], ignore_order)

  # --------------------------------------------------------------------------
  # STEP 2 - Eligibility criteria on the patients file
  # --------------------------------------------------------------------------
  # A) Two true Class I variants: both variants are in the true Class I list.
  true_c1 <- read_first_col(files$true_class1, ignore_order = ignore_order)
  two_true_c1 <- in_set(v1, true_c1) & in_set(v2, true_c1)

  # B) At least one F508del: either variant has F508del as a component
  #    (alone or inside a complex allele).
  f508 <- has_f508del(v1) | has_f508del(v2)

  # C) Variants designated responsive by EMA for ETI. The whole allele must
  #    appear in the list as written there (complex alleles included).
  ema_eti <- read_first_col(files$ema_eti, ignore_order = ignore_order)
  v1_ema_eti <- in_set(v1, ema_eti)
  v2_ema_eti <- in_set(v2, ema_eti)

  # D) Variants designated responsive by EMA for VTD (same rule as C).
  ema_vtd <- read_first_col(files$ema_vtd, ignore_order = ignore_order)
  v1_ema_vtd <- in_set(v1, ema_vtd)
  v2_ema_vtd <- in_set(v2, ema_vtd)

  # Annotated patients table: original columns + the criteria above.
  # "X" marks a patient that meets the criterion; the per-variant EMA columns
  # contain "yes"/"no".
  patients_out <- data.frame(
    patients,
    `TWO TRUE CLASS I variants (X)`                         = mark_x(two_true_c1),
    `≥1 F508del variant (X)`                                = mark_x(f508),
    `non-F508del VARIANT 1 EMA designated responsive for ETI (yes/no)`  = yes_no(v1_ema_eti),
    `non-F508del VARIANT 2 EMA designated responsive for ETI (yes/no)`  = yes_no(v2_ema_eti),
    `≥1 EMA non-F508del responsive variant for ETI (x)`                 = mark_x(v1_ema_eti | v2_ema_eti),
    `non-F508del VARIANT 1 EMA designated responsive for VTD (yes/no)`  = yes_no(v1_ema_vtd),
    `non-F508del VARIANT 2 EMA designated responsive for VTD (yes/no)`  = yes_no(v2_ema_vtd),
    `≥1 EMA non-F508del responsive variant for VTD (x)`                 = mark_x(v1_ema_vtd | v2_ema_vtd),
    check.names = FALSE, stringsAsFactors = FALSE
  )
  names(patients_out)[1:3] <- names(patients)

  # --------------------------------------------------------------------------
  # STEP 3 - Select the patients for the framework
  # --------------------------------------------------------------------------
  # A patient is kept only if NONE of the four summary criteria applies:
  # not two true Class I variants, no F508del, no EMA-responsive variant for
  # ETI, no EMA-responsive variant for VTD.
  keep <- !(two_true_c1 | f508 | v1_ema_eti | v2_ema_eti | v1_ema_vtd | v2_ema_vtd)

  fw_id <- patients[[1]][keep]
  fw_o1 <- patients[[2]][keep]            # variant names as in the patients file
  fw_o2 <- patients[[3]][keep]
  fw_v1 <- v1[keep]                       # matching keys of the same variants
  fw_v2 <- v2[keep]

  cat("Patients in PATIENTS_FILE :", nrow(patients), "\n")
  cat("Patients in FRAMEWORK     :", length(fw_id), "\n")

  # --------------------------------------------------------------------------
  # STEP 4 - Evidence for each variant of the selected patients
  # --------------------------------------------------------------------------
  # Every evidence is computed separately for variant 1 and variant 2.
  # Variants with no match in a list get an empty cell.

  # French cohort real-world data (RWD): a single column per variant that
  # collects three sources:
  #   - the category of the variant in the French cohort file (Responsive,
  #     Probably responsive, Probably non-responsive, Non-responsive,
  #     Inconclusive);
  #   - "True Class I" if the variant is in the true Class I list;
  #   - "Exceptional Class I" if the variant is in the exceptional Class I list.
  # If a variant is found in more than one source a warning is issued and the
  # more specific label wins (Exceptional Class I > True Class I > category).
  fcp <- read_category_lookup(files$fcp_rwd, ignore_order = ignore_order)
  rwd1 <- lookup_category(fw_v1, fcp)
  rwd2 <- lookup_category(fw_v2, fcp)
  exc_c1 <- read_first_col(files$exceptional_c1, ignore_order = ignore_order)
  add_label <- function(current, flag, label) {
    if (any(flag & !is.na(current)))
      warning("Variants with more than one RWD label: '", label, "' takes precedence.",
              call. = FALSE)
    current[flag] <- label
    current
  }
  rwd1 <- add_label(rwd1, in_set(fw_v1, true_c1), "True Class I")
  rwd2 <- add_label(rwd2, in_set(fw_v2, true_c1), "True Class I")
  rwd1 <- add_label(rwd1, in_set(fw_v1, exc_c1), "Exceptional Class I")
  rwd2 <- add_label(rwd2, in_set(fw_v2, exc_c1), "Exceptional Class I")

  # FDA lists: "X" if the variant is listed as responsive to ETI / VTD.
  fda_eti <- read_first_col(files$fda_eti, ignore_order = ignore_order)
  fda_vtd <- read_first_col(files$fda_vtd, ignore_order = ignore_order)

  # In vitro studies: category of the variant (Responsive, Borderline or
  # Non-responsive) for ETI and for VTI. These two files are used only for the
  # framework table; the annotated patients table does not use them.
  iv_eti <- read_in_vitro_lookup(files$in_vitro_eti, ignore_order = ignore_order)
  iv_vti <- read_in_vitro_lookup(files$in_vitro_vti, ignore_order = ignore_order)

  # --------------------------------------------------------------------------
  # Build the framework table from the template
  # --------------------------------------------------------------------------
  # Header row of the template (see FRAMEWORK TEMPLATE above).
  tpl_hdr <- names(read_excel(files$framework_template, n_max = 0,
                              .name_repair = "minimal"))
  n_fw <- length(fw_id)

  # Assigns each result to the template column recognized by its key words.
  # 'values' is indexed by column position (stored as text).
  values <- list()
  values[[as.character(find_template_col(tpl_hdr, "^patient\\s*id", what = "Patient ID"))]]                     <- fw_id
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*1", "legacy", what = "VARIANT 1 legacy name"))]]   <- fw_o1
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*2", "legacy", what = "VARIANT 2 legacy name"))]]   <- fw_o2
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*1", "RWD", what = "VARIANT 1 RWD"))]]              <- rwd1
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*2", "RWD", what = "VARIANT 2 RWD"))]]              <- rwd2
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*1", "FDA", "ETI", what = "VARIANT 1 ETI FDA"))]]     <- mark_x(in_set(fw_v1, fda_eti))
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*2", "FDA", "ETI", what = "VARIANT 2 ETI FDA"))]]     <- mark_x(in_set(fw_v2, fda_eti))
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*1", "FDA", "VTD", what = "VARIANT 1 VTD FDA"))]]     <- mark_x(in_set(fw_v1, fda_vtd))
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*2", "FDA", "VTD", what = "VARIANT 2 VTD FDA"))]]     <- mark_x(in_set(fw_v2, fda_vtd))
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*1", "in\\s*vitro", "ETI", what = "VARIANT 1 in vitro ETI"))]] <- lookup_category(fw_v1, iv_eti)
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*1", "in\\s*vitro", "^(?!.*ETI)", what = "VARIANT 1 in vitro VTI"))]] <- lookup_category(fw_v1, iv_vti)
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*2", "in\\s*vitro", "ETI", what = "VARIANT 2 in vitro ETI"))]] <- lookup_category(fw_v2, iv_eti)
  values[[as.character(find_template_col(tpl_hdr, "variant\\s*2", "in\\s*vitro", "^(?!.*ETI)", what = "VARIANT 2 in vitro VTI"))]] <- lookup_category(fw_v2, iv_vti)

  # Table with exactly the template columns, in the template order; columns
  # not recognized above stay empty.
  fw <- as.data.frame(matrix(NA_character_, nrow = n_fw, ncol = length(tpl_hdr)),
                      stringsAsFactors = FALSE)
  names(fw) <- tpl_hdr
  for (j in as.integer(names(values))) fw[[j]] <- values[[as.character(j)]]

  # --------------------------------------------------------------------------
  # Write the output files
  # --------------------------------------------------------------------------
  if (write_files) {
    # Annotated patients table, named after the input patients file
    # ("my_registry.xlsx" -> "my_registry_annotated.xlsx").
    pat_name <- paste0(tools::file_path_sans_ext(basename(files$patients)),
                       "_annotated.xlsx")
    write.xlsx(patients_out, file.path(output_dir, pat_name), overwrite = TRUE)

    # Framework: the template workbook itself is filled in from row 2 (row 1
    # is the header), so sheet name, header text and formatting are kept.
    # The file is saved in output_dir with the template's name.
    wb <- loadWorkbook(files$framework_template)
    old_rows <- nrow(read_excel(files$framework_template, col_types = "text"))
    if (old_rows > 0) {
      # A template that already contains data rows is cleared first.
      warning("The FRAMEWORK template already contains ", old_rows,
              " data row(s): they are removed before populating.", call. = FALSE)
      deleteData(wb, sheet = 1, cols = seq_along(tpl_hdr),
                 rows = 2:(old_rows + 1), gridExpand = TRUE)
    }
    writeData(wb, sheet = 1, x = fw, startRow = 2, colNames = FALSE)
    fw_path <- file.path(output_dir, basename(files$framework_template))
    # If output_dir is the template's own folder, "_populated" is appended to
    # the name so that the template is never overwritten.
    if (file.exists(files$framework_template) &&
        normalizePath(fw_path, mustWork = FALSE) ==
        normalizePath(files$framework_template))
      fw_path <- file.path(output_dir, paste0(
        tools::file_path_sans_ext(basename(files$framework_template)),
        "_populated.xlsx"))
    saveWorkbook(wb, fw_path, overwrite = TRUE)
    cat("\nFiles written to:", normalizePath(output_dir), "\n")
  }

  # Console summary: number of annotated cells in each evidence column.
  cat("\nEvidence summary (number of annotated allele rows):\n")
  for (nm in names(fw)[-(1:3)])
    cat(sprintf("  %-110s %d\n", nm, sum(!is.na(fw[[nm]]))))

  invisible(list(patients = patients_out, framework = fw))
}
