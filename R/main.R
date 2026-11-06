source("registry_evidence_integration.R")

  files <- list(
    patients       = "PATIENTS_FILE_TEST.xlsx",
    framework_template = "Registry_based_evidence_integrated_framework_output.xlsx",
    true_class1    = "ETI_real_world_data_true_class_I_FCP.xlsx",
    exceptional_c1 = "ETI_real_world_data_exceptional_class_I_FCP.xlsx",
    fcp_rwd        = "ETI_real_world_data_FCP.xlsx",
    ema_eti        = "ETI_responsive_variants_EMA.xlsx",
    ema_vtd        = "VTD_responsive_variants_EMA.xlsx",
    fda_eti        = "Variants_individually_listed_as_responsive_to_ETI_FDA.xlsx",
    fda_vtd        = "Variants_individually_listed_as_responsive_to_VTD_FDA.xlsx",
    in_vitro_eti   = "655_variants_tested_in_vitro_with_ETI.xlsx",
    in_vitro_vti   = "417_variants_tested_in_vitro_with_VTI.xlsx"
  )
  res <- run_registry_framework(files, output_dir = "output")
  